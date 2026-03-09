//! Chart XML parsing and serialization.
//!
//! This module provides functions to parse chart XML parts (`<c:chartSpace>`)
//! into `Chart` domain objects, serialize them back to XML, extract chart
//! references from drawing XML, and load charts during workbook open.

use std::io::Cursor;

use offidized_opc::relationship::TargetMode;
use offidized_opc::uri::PartUri;
use offidized_opc::{Package, Part};
use quick_xml::events::{BytesDecl, BytesEnd, BytesRef, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::chart::{
    BarDirection, Chart, ChartAxis, ChartDataRef, ChartGrouping, ChartLegend, ChartSeries,
    ChartType,
};
use crate::error::Result;
use crate::worksheet::Worksheet;

const DRAWING_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const CHART_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const DRAWINGML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELATIONSHIP_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn decode_general_ref(event: &BytesRef<'_>) -> Result<String> {
    let reference = event.decode().map_err(quick_xml::Error::from)?;
    let escaped = format!("&{};", reference);
    quick_xml::escape::unescape(&escaped)
        .map(|text| text.into_owned())
        .map_err(quick_xml::Error::from)
        .map_err(Into::into)
}

fn assign_string_cache_value(values: &mut Vec<String>, idx: usize, value: String) {
    if values.len() <= idx {
        values.resize(idx + 1, String::new());
    }
    values[idx] = value;
}

fn assign_numeric_cache_value(values: &mut Vec<Option<f64>>, idx: usize, value: &str) {
    if values.len() <= idx {
        values.resize(idx + 1, None);
    }
    values[idx] = value.parse().ok();
}

#[derive(Debug, Clone)]
struct ParsedDrawingChartRef {
    relationship_id: String,
    from_col: u32,
    from_row: u32,
    from_col_off: i64,
    from_row_off: i64,
    to_col: u32,
    to_row: u32,
    to_col_off: i64,
    to_row_off: i64,
    name: Option<String>,
    /// Extent width in EMUs (from `<ext cx="..."/>` in one-cell anchors).
    extent_cx: Option<i64>,
    /// Extent height in EMUs (from `<ext cy="..."/>` in one-cell anchors).
    extent_cy: Option<i64>,
}

/// Extracts chart references (with anchor data) from a drawing XML part.
///
/// Charts are referenced inside `<a:graphicData>` elements with a `<c:chart r:id="..."/>`
/// child. This function scans for both `twoCellAnchor` and `oneCellAnchor` elements that
/// contain chart graphic frames, and captures anchor positioning and chart name.
fn parse_drawing_chart_refs(xml: &[u8]) -> Result<Vec<ParsedDrawingChartRef>> {
    #[derive(Clone, Copy, PartialEq, Eq)]
    enum AnchorKind {
        OneCell,
        TwoCell,
    }

    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut charts = Vec::new();

    // Anchor tracking state
    let mut in_anchor = false;
    let mut anchor_kind: Option<AnchorKind> = None;
    let mut in_from = false;
    let mut in_to = false;
    let mut in_xfrm = false;
    let mut in_col = false;
    let mut in_col_off = false;
    let mut in_row = false;
    let mut in_row_off = false;

    let mut from_col: u32 = 0;
    let mut from_row: u32 = 0;
    let mut from_col_off: i64 = 0;
    let mut from_row_off: i64 = 0;
    let mut to_col: u32 = 0;
    let mut to_row: u32 = 0;
    let mut to_col_off: i64 = 0;
    let mut to_row_off: i64 = 0;
    let mut chart_name: Option<String> = None;
    let mut extent_cx: Option<i64> = None;
    let mut extent_cy: Option<i64> = None;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                match local {
                    b"twoCellAnchor" => {
                        in_anchor = true;
                        anchor_kind = Some(AnchorKind::TwoCell);
                        from_col = 0;
                        from_row = 0;
                        from_col_off = 0;
                        from_row_off = 0;
                        to_col = 0;
                        to_row = 0;
                        to_col_off = 0;
                        to_row_off = 0;
                        chart_name = None;
                        extent_cx = None;
                        extent_cy = None;
                    }
                    b"oneCellAnchor" => {
                        in_anchor = true;
                        anchor_kind = Some(AnchorKind::OneCell);
                        from_col = 0;
                        from_row = 0;
                        from_col_off = 0;
                        from_row_off = 0;
                        to_col = 0;
                        to_row = 0;
                        to_col_off = 0;
                        to_row_off = 0;
                        chart_name = None;
                        extent_cx = None;
                        extent_cy = None;
                    }
                    b"from" if in_anchor => {
                        in_from = true;
                    }
                    b"to" if in_anchor => {
                        in_to = true;
                    }
                    b"xfrm" if in_anchor => {
                        in_xfrm = true;
                    }
                    b"col" if in_from || in_to => {
                        in_col = true;
                    }
                    b"colOff" if in_from || in_to => {
                        in_col_off = true;
                    }
                    b"row" if in_from || in_to => {
                        in_row = true;
                    }
                    b"rowOff" if in_from || in_to => {
                        in_row_off = true;
                    }
                    _ => {}
                }

                // Also check for <c:chart> in Start events (rare but possible)
                if local == b"chart" && in_anchor {
                    if let Some(chart_ref) = extract_chart_ref_from_event(
                        event,
                        from_col,
                        from_row,
                        from_col_off,
                        from_row_off,
                        to_col,
                        to_row,
                        to_col_off,
                        to_row_off,
                        &chart_name,
                        extent_cx,
                        extent_cy,
                    ) {
                        charts.push(chart_ref);
                    }
                }
                // Extract chart name from <xdr:cNvPr name="..."/>
                if local == b"cNvPr" && in_anchor {
                    for attribute in event.attributes().flatten() {
                        let attr_local = local_name(attribute.key.as_ref());
                        if attr_local == b"name" {
                            chart_name = Some(
                                String::from_utf8_lossy(attribute.value.as_ref()).into_owned(),
                            );
                        }
                    }
                }
            }
            Event::Empty(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                if local == b"chart" && in_anchor {
                    if let Some(chart_ref) = extract_chart_ref_from_event(
                        event,
                        from_col,
                        from_row,
                        from_col_off,
                        from_row_off,
                        to_col,
                        to_row,
                        to_col_off,
                        to_row_off,
                        &chart_name,
                        extent_cx,
                        extent_cy,
                    ) {
                        charts.push(chart_ref);
                    }
                }
                // Extract chart name from empty <xdr:cNvPr name="..."/>
                if local == b"cNvPr" && in_anchor {
                    for attribute in event.attributes().flatten() {
                        let attr_local = local_name(attribute.key.as_ref());
                        if attr_local == b"name" {
                            chart_name = Some(
                                String::from_utf8_lossy(attribute.value.as_ref()).into_owned(),
                            );
                        }
                    }
                }
                // Parse <ext cx="..." cy="..."/> for one-cell anchor sizing
                if local == b"ext"
                    && in_anchor
                    && anchor_kind == Some(AnchorKind::OneCell)
                    && !in_from
                    && !in_to
                    && !in_xfrm
                {
                    for attribute in event.attributes().flatten() {
                        let attr_local = local_name(attribute.key.as_ref());
                        let val_str = String::from_utf8_lossy(attribute.value.as_ref());
                        if attr_local == b"cx" {
                            extent_cx = val_str.parse::<i64>().ok();
                        } else if attr_local == b"cy" {
                            extent_cy = val_str.parse::<i64>().ok();
                        }
                    }
                }
            }
            Event::End(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                match local {
                    b"twoCellAnchor" | b"oneCellAnchor" => {
                        in_anchor = false;
                        anchor_kind = None;
                        in_from = false;
                        in_to = false;
                        in_xfrm = false;
                    }
                    b"from" => {
                        in_from = false;
                    }
                    b"to" => {
                        in_to = false;
                    }
                    b"xfrm" => {
                        in_xfrm = false;
                    }
                    b"col" => {
                        in_col = false;
                    }
                    b"colOff" => {
                        in_col_off = false;
                    }
                    b"row" => {
                        in_row = false;
                    }
                    b"rowOff" => {
                        in_row_off = false;
                    }
                    _ => {}
                }
            }
            Event::Text(ref text) => {
                let text_val = text.xml_content().unwrap_or_default();
                if in_col {
                    let val = text_val.parse::<u32>().unwrap_or(0);
                    if in_from {
                        from_col = val;
                    } else if in_to {
                        to_col = val;
                    }
                } else if in_col_off {
                    let val = text_val.parse::<i64>().unwrap_or(0);
                    if in_from {
                        from_col_off = val;
                    } else if in_to {
                        to_col_off = val;
                    }
                } else if in_row {
                    let val = text_val.parse::<u32>().unwrap_or(0);
                    if in_from {
                        from_row = val;
                    } else if in_to {
                        to_row = val;
                    }
                } else if in_row_off {
                    let val = text_val.parse::<i64>().unwrap_or(0);
                    if in_from {
                        from_row_off = val;
                    } else if in_to {
                        to_row_off = val;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    Ok(charts)
}

/// Helper to extract a chart ref from a `<c:chart r:id="..."/>` element.
#[allow(clippy::too_many_arguments)]
fn extract_chart_ref_from_event(
    event: &BytesStart<'_>,
    from_col: u32,
    from_row: u32,
    from_col_off: i64,
    from_row_off: i64,
    to_col: u32,
    to_row: u32,
    to_col_off: i64,
    to_row_off: i64,
    chart_name: &Option<String>,
    extent_cx: Option<i64>,
    extent_cy: Option<i64>,
) -> Option<ParsedDrawingChartRef> {
    for attribute in event.attributes().flatten() {
        let attr_local = local_name(attribute.key.as_ref());
        if attr_local == b"id" {
            let value = String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
            if !value.is_empty() {
                return Some(ParsedDrawingChartRef {
                    relationship_id: value,
                    from_col,
                    from_row,
                    from_col_off,
                    from_row_off,
                    to_col,
                    to_row,
                    to_col_off,
                    to_row_off,
                    name: chart_name.clone(),
                    extent_cx,
                    extent_cy,
                });
            }
        }
    }
    None
}

/// Parses a chart XML part (`<c:chartSpace>`) into a `Chart` domain object.
pub(crate) fn parse_chart_xml(xml: &[u8]) -> Result<Chart> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    let mut chart_type: Option<ChartType> = None;
    let mut bar_direction: Option<BarDirection> = None;
    let mut grouping: Option<ChartGrouping> = None;
    let mut series_list: Vec<ChartSeries> = Vec::new();
    let mut axes: Vec<ChartAxis> = Vec::new();
    let mut legend: Option<ChartLegend> = None;
    let mut chart_title: Option<String> = None;
    let mut vary_colors = false;

    // Nested state tracking
    let mut in_chart = false;
    let mut in_plot_area = false;
    let mut in_chart_type_element = false;
    let mut in_series = false;
    let mut in_legend = false;
    let mut in_title = false;
    let mut in_axis = false;
    let mut axis_type_name = String::new();

    // Series parsing state
    let mut ser_idx: u32 = 0;
    let mut ser_order: u32 = 0;
    let mut ser_name: Option<String> = None;
    let mut ser_name_ref: Option<String> = None;
    let mut ser_cat_formula: Option<String> = None;
    let mut ser_val_formula: Option<String> = None;
    let mut ser_cat_str_values: Vec<String> = Vec::new();
    let mut ser_cat_num_values: Vec<Option<f64>> = Vec::new();
    let mut ser_val_num_values: Vec<Option<f64>> = Vec::new();
    let mut current_formula = String::new();
    let mut current_cache_value = String::new();

    // Track which sub-element of <c:ser> we are inside
    let mut in_tx = false;
    let mut in_cat = false;
    let mut in_val = false;
    let mut in_str_ref = false;
    let mut in_num_ref = false;
    let mut in_str_cache = false;
    let mut in_num_cache = false;
    let mut in_pt = false;
    let mut current_pt_idx: usize = 0;
    let mut in_pt_v = false;
    let mut in_f = false;
    let mut in_tx_v = false;

    // Axis parsing state
    let mut ax_id: u32 = 0;
    let mut ax_position = String::new();
    let mut ax_title: Option<String> = None;
    let mut ax_min: Option<f64> = None;
    let mut ax_max: Option<f64> = None;
    let mut ax_major_gridlines = false;
    let mut ax_minor_gridlines = false;
    let mut ax_crosses: Option<u32> = None;
    let mut ax_deleted = false;
    let mut in_axis_title = false;
    let mut in_axis_scaling = false;

    // Title text accumulation
    let mut title_text_parts: Vec<String> = Vec::new();
    let mut in_title_rich = false;
    let mut in_title_p = false;
    let mut in_title_r = false;
    let mut in_title_t = false;

    // Axis title text accumulation
    let mut axis_title_text_parts: Vec<String> = Vec::new();
    let mut in_axis_title_rich = false;
    let mut in_axis_title_p = false;
    let mut in_axis_title_r = false;
    let mut in_axis_title_t = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                match local {
                    b"chart" if !in_chart => {
                        in_chart = true;
                    }
                    b"plotArea" if in_chart => {
                        in_plot_area = true;
                    }
                    b"title" if in_axis => {
                        in_axis_title = true;
                        axis_title_text_parts.clear();
                    }
                    b"title" if in_chart && !in_plot_area => {
                        in_title = true;
                        title_text_parts.clear();
                    }
                    b"rich" if in_axis_title => {
                        in_axis_title_rich = true;
                    }
                    b"rich" if in_title => {
                        in_title_rich = true;
                    }
                    b"p" if in_axis_title_rich => {
                        in_axis_title_p = true;
                    }
                    b"p" if in_title_rich => {
                        in_title_p = true;
                    }
                    b"r" if in_axis_title_p => {
                        in_axis_title_r = true;
                    }
                    b"r" if in_title_p => {
                        in_title_r = true;
                    }
                    b"t" if in_axis_title_r => {
                        in_axis_title_t = true;
                    }
                    b"t" if in_title_r => {
                        in_title_t = true;
                    }
                    b"legend" if in_chart => {
                        in_legend = true;
                    }
                    _ if in_plot_area && !in_chart_type_element && !in_axis => {
                        // Check if this is a chart type element
                        let local_str = std::str::from_utf8(local).unwrap_or("");
                        if let Some(ct) = ChartType::from_xml_value(local_str) {
                            chart_type = Some(ct);
                            in_chart_type_element = true;
                        }
                        // Check if this is an axis element
                        match local {
                            b"catAx" | b"valAx" | b"dateAx" | b"serAx" => {
                                in_axis = true;
                                axis_type_name =
                                    std::str::from_utf8(local).unwrap_or("").to_string();
                                ax_id = 0;
                                ax_position.clear();
                                ax_title = None;
                                ax_min = None;
                                ax_max = None;
                                ax_major_gridlines = false;
                                ax_minor_gridlines = false;
                                ax_crosses = None;
                                ax_deleted = false;
                            }
                            _ => {}
                        }
                    }
                    b"ser" if in_chart_type_element => {
                        in_series = true;
                        ser_idx = 0;
                        ser_order = 0;
                        ser_name = None;
                        ser_name_ref = None;
                        ser_cat_formula = None;
                        ser_val_formula = None;
                        ser_cat_str_values.clear();
                        ser_cat_num_values.clear();
                        ser_val_num_values.clear();
                    }
                    b"tx" if in_series => {
                        in_tx = true;
                    }
                    b"cat" if in_series => {
                        in_cat = true;
                    }
                    b"val" if in_series => {
                        in_val = true;
                    }
                    b"strRef" if in_tx || in_cat => {
                        in_str_ref = true;
                    }
                    b"numRef" if in_val || in_cat => {
                        in_num_ref = true;
                    }
                    b"f" if in_str_ref || in_num_ref => {
                        in_f = true;
                        current_formula.clear();
                    }
                    b"strCache" if in_str_ref => {
                        in_str_cache = true;
                    }
                    b"numCache" if in_num_ref => {
                        in_num_cache = true;
                    }
                    b"pt" if in_str_cache || in_num_cache => {
                        in_pt = true;
                        current_pt_idx = 0;
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"idx" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                current_pt_idx = value.trim().parse().unwrap_or(0);
                            }
                        }
                    }
                    b"v" if in_tx && !in_str_ref => {
                        in_tx_v = true;
                    }
                    b"v" if in_pt => {
                        in_pt_v = true;
                        current_cache_value.clear();
                    }
                    b"scaling" if in_axis => {
                        in_axis_scaling = true;
                    }
                    _ => {}
                }
            }
            Event::Empty(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                match local {
                    b"grouping" if in_chart_type_element => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                grouping = ChartGrouping::from_xml_value(value.trim());
                            }
                        }
                    }
                    b"barDir" if in_chart_type_element => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                bar_direction = BarDirection::from_xml_value(value.trim());
                            }
                        }
                    }
                    b"varyColors" if in_chart_type_element => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                vary_colors = value.trim() == "1" || value.trim() == "true";
                            }
                        }
                    }
                    b"idx" if in_series => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                ser_idx = value.trim().parse().unwrap_or(0);
                            }
                        }
                    }
                    b"order" if in_series => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                ser_order = value.trim().parse().unwrap_or(0);
                            }
                        }
                    }
                    b"legendPos" if in_legend => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                let mut leg = legend.take().unwrap_or_default();
                                leg.set_position(value.trim());
                                legend = Some(leg);
                            }
                        }
                    }
                    b"overlay" if in_legend => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                let overlay = value.trim() == "1" || value.trim() == "true";
                                let mut leg = legend.take().unwrap_or_default();
                                leg.set_overlay(overlay);
                                legend = Some(leg);
                            }
                        }
                    }
                    b"axId" if in_chart_type_element && !in_series => {
                        // axId references inside chart type elements (skip for now)
                    }
                    b"axId" if in_axis => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                ax_id = value.trim().parse().unwrap_or(0);
                            }
                        }
                    }
                    b"axPos" if in_axis => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                ax_position = String::from_utf8_lossy(attribute.value.as_ref())
                                    .trim()
                                    .to_string();
                            }
                        }
                    }
                    b"crossAx" if in_axis => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                ax_crosses = value.trim().parse().ok();
                            }
                        }
                    }
                    b"delete" if in_axis => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                ax_deleted = value.trim() == "1" || value.trim() == "true";
                            }
                        }
                    }
                    b"min" if in_axis_scaling => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                ax_min = value.trim().parse().ok();
                            }
                        }
                    }
                    b"max" if in_axis_scaling => {
                        for attribute in event.attributes().flatten() {
                            if local_name(attribute.key.as_ref()) == b"val" {
                                let value =
                                    String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                                ax_max = value.trim().parse().ok();
                            }
                        }
                    }
                    b"majorGridlines" if in_axis => {
                        ax_major_gridlines = true;
                    }
                    b"minorGridlines" if in_axis => {
                        ax_minor_gridlines = true;
                    }
                    _ => {}
                }
            }
            Event::Text(ref event) => {
                let text = event
                    .xml_content()
                    .map_err(quick_xml::Error::from)?
                    .into_owned();
                if in_f {
                    current_formula.push_str(&text);
                }
                if in_pt_v {
                    current_cache_value.push_str(&text);
                }
                if in_tx_v {
                    ser_name = Some(text.to_string());
                }
                if in_axis_title_t {
                    axis_title_text_parts.push(text.to_string());
                } else if in_title_t {
                    title_text_parts.push(text.to_string());
                }
            }
            Event::GeneralRef(ref event) => {
                if in_f {
                    current_formula.push_str(&decode_general_ref(event)?);
                }
                if in_pt_v {
                    current_cache_value.push_str(&decode_general_ref(event)?);
                }
            }
            Event::End(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                match local {
                    b"chart" => {
                        in_chart = false;
                    }
                    b"plotArea" => {
                        in_plot_area = false;
                    }
                    b"title" if in_axis_title => {
                        in_axis_title = false;
                        if !axis_title_text_parts.is_empty() {
                            ax_title = Some(axis_title_text_parts.join(""));
                        }
                        in_axis_title_rich = false;
                        in_axis_title_p = false;
                        in_axis_title_r = false;
                        in_axis_title_t = false;
                    }
                    b"title" if in_title => {
                        in_title = false;
                        if !title_text_parts.is_empty() {
                            chart_title = Some(title_text_parts.join(""));
                        }
                        in_title_rich = false;
                        in_title_p = false;
                        in_title_r = false;
                        in_title_t = false;
                    }
                    b"rich" if in_axis_title_rich => {
                        in_axis_title_rich = false;
                    }
                    b"rich" if in_title_rich => {
                        in_title_rich = false;
                    }
                    b"p" if in_axis_title_p => {
                        in_axis_title_p = false;
                    }
                    b"p" if in_title_p => {
                        in_title_p = false;
                    }
                    b"r" if in_axis_title_r => {
                        in_axis_title_r = false;
                    }
                    b"r" if in_title_r => {
                        in_title_r = false;
                    }
                    b"t" if in_axis_title_t => {
                        in_axis_title_t = false;
                    }
                    b"t" if in_title_t => {
                        in_title_t = false;
                    }
                    b"legend" => {
                        in_legend = false;
                    }
                    b"scaling" => {
                        in_axis_scaling = false;
                    }
                    b"f" => {
                        let trimmed = current_formula.trim().to_string();
                        if !trimmed.is_empty() {
                            if in_tx && in_str_ref {
                                ser_name_ref = Some(trimmed);
                            } else if in_cat && (in_str_ref || in_num_ref) {
                                ser_cat_formula = Some(trimmed);
                            } else if in_val && in_num_ref {
                                ser_val_formula = Some(trimmed);
                            }
                        }
                        in_f = false;
                        current_formula.clear();
                    }
                    b"strCache" => {
                        in_str_cache = false;
                    }
                    b"numCache" => {
                        in_num_cache = false;
                    }
                    b"pt" => {
                        in_pt = false;
                    }
                    b"v" if in_pt_v => {
                        let trimmed = current_cache_value.trim().to_string();
                        if in_cat {
                            if in_str_cache {
                                assign_string_cache_value(
                                    &mut ser_cat_str_values,
                                    current_pt_idx,
                                    trimmed,
                                );
                            } else if in_num_cache {
                                assign_numeric_cache_value(
                                    &mut ser_cat_num_values,
                                    current_pt_idx,
                                    &trimmed,
                                );
                            }
                        } else if in_val && in_num_cache {
                            assign_numeric_cache_value(
                                &mut ser_val_num_values,
                                current_pt_idx,
                                &trimmed,
                            );
                        }
                        in_pt_v = false;
                        current_cache_value.clear();
                    }
                    b"strRef" => {
                        in_str_ref = false;
                    }
                    b"numRef" => {
                        in_num_ref = false;
                    }
                    b"tx" => {
                        in_tx = false;
                    }
                    b"v" => {
                        in_tx_v = false;
                    }
                    b"cat" => {
                        in_cat = false;
                    }
                    b"val" => {
                        in_val = false;
                    }
                    b"ser" => {
                        if in_series {
                            let mut series = ChartSeries::new(ser_idx, ser_order);
                            if let Some(ref name) = ser_name {
                                series.set_name(name.as_str());
                            }
                            if let Some(ref name_ref) = ser_name_ref {
                                series.set_name_ref(name_ref.as_str());
                            }
                            if ser_cat_formula.is_some()
                                || !ser_cat_str_values.is_empty()
                                || !ser_cat_num_values.is_empty()
                            {
                                let mut categories = ser_cat_formula
                                    .as_deref()
                                    .map(ChartDataRef::from_formula)
                                    .unwrap_or_default();
                                if !ser_cat_str_values.is_empty() {
                                    categories.set_str_values(ser_cat_str_values.clone());
                                }
                                if !ser_cat_num_values.is_empty() {
                                    categories.set_num_values(ser_cat_num_values.clone());
                                }
                                series.set_categories(categories);
                            }
                            if ser_val_formula.is_some() || !ser_val_num_values.is_empty() {
                                let mut values = ser_val_formula
                                    .as_deref()
                                    .map(ChartDataRef::from_formula)
                                    .unwrap_or_default();
                                if !ser_val_num_values.is_empty() {
                                    values.set_num_values(ser_val_num_values.clone());
                                }
                                series.set_values(values);
                            }
                            series_list.push(series);
                        }
                        in_series = false;
                        in_tx = false;
                        in_cat = false;
                        in_val = false;
                        in_str_ref = false;
                        in_num_ref = false;
                        in_str_cache = false;
                        in_num_cache = false;
                        in_pt = false;
                        in_pt_v = false;
                        in_f = false;
                        in_tx_v = false;
                    }
                    b"catAx" | b"valAx" | b"dateAx" | b"serAx" => {
                        if in_axis {
                            let mut axis = ChartAxis::new_category();
                            axis.set_id(ax_id);
                            axis.set_axis_type(axis_type_name.as_str());
                            if !ax_position.is_empty() {
                                axis.set_position(ax_position.as_str());
                            }
                            if let Some(ref title) = ax_title {
                                axis.set_title(title.as_str());
                            }
                            if let Some(min) = ax_min {
                                axis.set_min(min);
                            }
                            if let Some(max) = ax_max {
                                axis.set_max(max);
                            }
                            axis.set_major_gridlines(ax_major_gridlines);
                            axis.set_minor_gridlines(ax_minor_gridlines);
                            if let Some(crosses) = ax_crosses {
                                axis.set_crosses_ax(crosses);
                            } else {
                                axis.clear_crosses_ax();
                            }
                            axis.set_deleted(ax_deleted);
                            axes.push(axis);
                        }
                        in_axis = false;
                        in_axis_title = false;
                        in_axis_title_rich = false;
                        in_axis_title_p = false;
                        in_axis_title_r = false;
                        in_axis_title_t = false;
                        in_axis_scaling = false;
                    }
                    _ if in_plot_area && in_chart_type_element && !in_series && !in_axis => {
                        // Closing a chart type element (e.g. </c:barChart>)
                        let local_str = std::str::from_utf8(local).unwrap_or("");
                        if ChartType::from_xml_value(local_str).is_some() {
                            in_chart_type_element = false;
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    let ct = chart_type.unwrap_or(ChartType::Bar);
    let mut chart = Chart::new(ct);
    // Override defaults with what was actually in the XML
    if let Some(title) = chart_title {
        chart.set_title(title);
    }
    if let Some(dir) = bar_direction {
        chart.set_bar_direction(dir);
    }
    if let Some(grp) = grouping {
        chart.set_grouping(grp);
    }
    chart.set_vary_colors(vary_colors);
    for s in series_list {
        chart.add_series(s);
    }
    // Replace default axes with parsed axes (XML is the source of truth)
    chart.clear_axes();
    for a in axes {
        chart.add_axis(a);
    }
    if let Some(leg) = legend {
        chart.set_legend(leg);
    }

    Ok(chart)
}

/// Serializes a `Chart` domain object into chart XML bytes (`<c:chartSpace>`).
pub(crate) fn serialize_chart_xml(chart: &Chart) -> Result<Vec<u8>> {
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut root = BytesStart::new("c:chartSpace");
    root.push_attribute(("xmlns:c", CHART_NS));
    root.push_attribute(("xmlns:a", DRAWINGML_NS));
    root.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    writer.write_event(Event::Start(root))?;

    writer.write_event(Event::Start(BytesStart::new("c:chart")))?;

    // Title
    if let Some(title) = chart.title() {
        writer.write_event(Event::Start(BytesStart::new("c:title")))?;
        writer.write_event(Event::Start(BytesStart::new("c:tx")))?;
        writer.write_event(Event::Start(BytesStart::new("c:rich")))?;
        writer.write_event(Event::Empty(BytesStart::new("a:bodyPr")))?;
        writer.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;
        writer.write_event(Event::Start(BytesStart::new("a:p")))?;
        writer.write_event(Event::Start(BytesStart::new("a:r")))?;
        writer.write_event(Event::Start(BytesStart::new("a:t")))?;
        writer.write_event(Event::Text(BytesText::new(title)))?;
        writer.write_event(Event::End(BytesEnd::new("a:t")))?;
        writer.write_event(Event::End(BytesEnd::new("a:r")))?;
        writer.write_event(Event::End(BytesEnd::new("a:p")))?;
        writer.write_event(Event::End(BytesEnd::new("c:rich")))?;
        writer.write_event(Event::End(BytesEnd::new("c:tx")))?;
        writer.write_event(Event::End(BytesEnd::new("c:title")))?;
    }

    // Plot area
    writer.write_event(Event::Start(BytesStart::new("c:plotArea")))?;
    writer.write_event(Event::Empty(BytesStart::new("c:layout")))?;

    // Chart type element
    let chart_type_tag = chart.chart_type().as_str();
    let chart_type_tag_c = format!("c:{chart_type_tag}");
    writer.write_event(Event::Start(BytesStart::new(chart_type_tag_c.as_str())))?;

    // Bar direction (must precede grouping per OOXML spec)
    if let Some(dir) = chart.bar_direction() {
        let mut bar_dir_tag = BytesStart::new("c:barDir");
        bar_dir_tag.push_attribute(("val", dir.as_str()));
        writer.write_event(Event::Empty(bar_dir_tag))?;
    }

    // Grouping
    if let Some(grp) = chart.grouping() {
        let mut grouping_tag = BytesStart::new("c:grouping");
        grouping_tag.push_attribute(("val", grp.as_str()));
        writer.write_event(Event::Empty(grouping_tag))?;
    }

    // Vary colors
    if chart.vary_colors() {
        let mut vc = BytesStart::new("c:varyColors");
        vc.push_attribute(("val", "1"));
        writer.write_event(Event::Empty(vc))?;
    }

    // Series
    for series in chart.series() {
        writer.write_event(Event::Start(BytesStart::new("c:ser")))?;

        let idx_text = series.idx().to_string();
        let mut idx_tag = BytesStart::new("c:idx");
        idx_tag.push_attribute(("val", idx_text.as_str()));
        writer.write_event(Event::Empty(idx_tag))?;

        let order_text = series.order().to_string();
        let mut order_tag = BytesStart::new("c:order");
        order_tag.push_attribute(("val", order_text.as_str()));
        writer.write_event(Event::Empty(order_tag))?;

        // For line/radar/area charts, suppress markers by default so
        // LibreOffice (and Excel) render connected lines, not dots.
        if matches!(
            chart.chart_type(),
            ChartType::Line | ChartType::Radar | ChartType::Area
        ) {
            writer.write_event(Event::Start(BytesStart::new("c:marker")))?;
            let mut symbol = BytesStart::new("c:symbol");
            symbol.push_attribute(("val", "none"));
            writer.write_event(Event::Empty(symbol))?;
            writer.write_event(Event::End(BytesEnd::new("c:marker")))?;
        }

        // Series name reference / literal
        if let Some(name_ref) = series.name_ref() {
            writer.write_event(Event::Start(BytesStart::new("c:tx")))?;
            writer.write_event(Event::Start(BytesStart::new("c:strRef")))?;
            writer.write_event(Event::Start(BytesStart::new("c:f")))?;
            writer.write_event(Event::Text(BytesText::new(name_ref)))?;
            writer.write_event(Event::End(BytesEnd::new("c:f")))?;
            writer.write_event(Event::End(BytesEnd::new("c:strRef")))?;
            writer.write_event(Event::End(BytesEnd::new("c:tx")))?;
        } else if let Some(name) = series.name() {
            writer.write_event(Event::Start(BytesStart::new("c:tx")))?;
            writer.write_event(Event::Start(BytesStart::new("c:v")))?;
            writer.write_event(Event::Text(BytesText::new(name)))?;
            writer.write_event(Event::End(BytesEnd::new("c:v")))?;
            writer.write_event(Event::End(BytesEnd::new("c:tx")))?;
        }

        // Categories
        if let Some(categories) = series.categories() {
            if categories.formula().is_some() || !categories.str_values().is_empty() {
                writer.write_event(Event::Start(BytesStart::new("c:cat")))?;
                writer.write_event(Event::Start(BytesStart::new("c:strRef")))?;
                if let Some(formula) = categories.formula() {
                    writer.write_event(Event::Start(BytesStart::new("c:f")))?;
                    writer.write_event(Event::Text(BytesText::new(formula)))?;
                    writer.write_event(Event::End(BytesEnd::new("c:f")))?;
                }
                if !categories.str_values().is_empty() {
                    write_str_cache(&mut writer, categories.str_values())?;
                }
                writer.write_event(Event::End(BytesEnd::new("c:strRef")))?;
                writer.write_event(Event::End(BytesEnd::new("c:cat")))?;
            }
        }

        // Values
        if let Some(values) = series.values() {
            if values.formula().is_some() || !values.num_values().is_empty() {
                writer.write_event(Event::Start(BytesStart::new("c:val")))?;
                writer.write_event(Event::Start(BytesStart::new("c:numRef")))?;
                if let Some(formula) = values.formula() {
                    writer.write_event(Event::Start(BytesStart::new("c:f")))?;
                    writer.write_event(Event::Text(BytesText::new(formula)))?;
                    writer.write_event(Event::End(BytesEnd::new("c:f")))?;
                }
                if !values.num_values().is_empty() {
                    write_num_cache(&mut writer, values.num_values())?;
                }
                writer.write_event(Event::End(BytesEnd::new("c:numRef")))?;
                writer.write_event(Event::End(BytesEnd::new("c:val")))?;
            }
        }

        // Explicit smooth=false for line charts (LibreOffice needs this)
        if chart.chart_type() == ChartType::Line {
            let mut smooth = BytesStart::new("c:smooth");
            smooth.push_attribute(("val", "0"));
            writer.write_event(Event::Empty(smooth))?;
        }

        writer.write_event(Event::End(BytesEnd::new("c:ser")))?;
    }

    // Axis ID references inside the chart type element
    for axis in chart.axes() {
        let ax_id_text = axis.id().to_string();
        let mut ax_id_tag = BytesStart::new("c:axId");
        ax_id_tag.push_attribute(("val", ax_id_text.as_str()));
        writer.write_event(Event::Empty(ax_id_tag))?;
    }

    writer.write_event(Event::End(BytesEnd::new(chart_type_tag_c.as_str())))?;

    // Axes
    for axis in chart.axes() {
        let axis_tag = format!("c:{}", axis.axis_type());
        writer.write_event(Event::Start(BytesStart::new(axis_tag.as_str())))?;

        let ax_id_text = axis.id().to_string();
        let mut ax_id_tag = BytesStart::new("c:axId");
        ax_id_tag.push_attribute(("val", ax_id_text.as_str()));
        writer.write_event(Event::Empty(ax_id_tag))?;

        writer.write_event(Event::Start(BytesStart::new("c:scaling")))?;
        // orientation is required by LibreOffice for correct axis rendering
        let mut orientation = BytesStart::new("c:orientation");
        orientation.push_attribute(("val", "minMax"));
        writer.write_event(Event::Empty(orientation))?;
        if let Some(min) = axis.min() {
            let min_text = min.to_string();
            let mut min_tag = BytesStart::new("c:min");
            min_tag.push_attribute(("val", min_text.as_str()));
            writer.write_event(Event::Empty(min_tag))?;
        }
        if let Some(max) = axis.max() {
            let max_text = max.to_string();
            let mut max_tag = BytesStart::new("c:max");
            max_tag.push_attribute(("val", max_text.as_str()));
            writer.write_event(Event::Empty(max_tag))?;
        }
        writer.write_event(Event::End(BytesEnd::new("c:scaling")))?;

        if axis.deleted() {
            let mut del = BytesStart::new("c:delete");
            del.push_attribute(("val", "1"));
            writer.write_event(Event::Empty(del))?;
        }

        let mut ax_pos = BytesStart::new("c:axPos");
        ax_pos.push_attribute(("val", axis.position()));
        writer.write_event(Event::Empty(ax_pos))?;

        if axis.major_gridlines() {
            writer.write_event(Event::Empty(BytesStart::new("c:majorGridlines")))?;
        }
        if axis.minor_gridlines() {
            writer.write_event(Event::Empty(BytesStart::new("c:minorGridlines")))?;
        }

        // Axis title
        if let Some(title) = axis.title() {
            writer.write_event(Event::Start(BytesStart::new("c:title")))?;
            writer.write_event(Event::Start(BytesStart::new("c:tx")))?;
            writer.write_event(Event::Start(BytesStart::new("c:rich")))?;
            writer.write_event(Event::Empty(BytesStart::new("a:bodyPr")))?;
            writer.write_event(Event::Empty(BytesStart::new("a:lstStyle")))?;
            writer.write_event(Event::Start(BytesStart::new("a:p")))?;
            writer.write_event(Event::Start(BytesStart::new("a:r")))?;
            writer.write_event(Event::Start(BytesStart::new("a:t")))?;
            writer.write_event(Event::Text(BytesText::new(title)))?;
            writer.write_event(Event::End(BytesEnd::new("a:t")))?;
            writer.write_event(Event::End(BytesEnd::new("a:r")))?;
            writer.write_event(Event::End(BytesEnd::new("a:p")))?;
            writer.write_event(Event::End(BytesEnd::new("c:rich")))?;
            writer.write_event(Event::End(BytesEnd::new("c:tx")))?;
            writer.write_event(Event::End(BytesEnd::new("c:title")))?;
        }

        if let Some(cross_ax) = axis.crosses_ax() {
            let cross_text = cross_ax.to_string();
            let mut cross_tag = BytesStart::new("c:crossAx");
            cross_tag.push_attribute(("val", cross_text.as_str()));
            writer.write_event(Event::Empty(cross_tag))?;
        }

        writer.write_event(Event::End(BytesEnd::new(axis_tag.as_str())))?;
    }

    writer.write_event(Event::End(BytesEnd::new("c:plotArea")))?;

    // Legend
    if let Some(legend) = chart.legend() {
        writer.write_event(Event::Start(BytesStart::new("c:legend")))?;
        let mut legend_pos = BytesStart::new("c:legendPos");
        legend_pos.push_attribute(("val", legend.position()));
        writer.write_event(Event::Empty(legend_pos))?;
        if legend.overlay() {
            let mut overlay_tag = BytesStart::new("c:overlay");
            overlay_tag.push_attribute(("val", "1"));
            writer.write_event(Event::Empty(overlay_tag))?;
        }
        writer.write_event(Event::End(BytesEnd::new("c:legend")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("c:chart")))?;
    writer.write_event(Event::End(BytesEnd::new("c:chartSpace")))?;

    Ok(writer.into_inner())
}

fn write_str_cache(writer: &mut Writer<Vec<u8>>, values: &[String]) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("c:strCache")))?;

    let count_text = values.len().to_string();
    let mut pt_count = BytesStart::new("c:ptCount");
    pt_count.push_attribute(("val", count_text.as_str()));
    writer.write_event(Event::Empty(pt_count))?;

    for (idx, value) in values.iter().enumerate() {
        let idx_text = idx.to_string();
        let mut pt = BytesStart::new("c:pt");
        pt.push_attribute(("idx", idx_text.as_str()));
        writer.write_event(Event::Start(pt))?;
        writer.write_event(Event::Start(BytesStart::new("c:v")))?;
        writer.write_event(Event::Text(BytesText::new(value.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("c:v")))?;
        writer.write_event(Event::End(BytesEnd::new("c:pt")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("c:strCache")))?;
    Ok(())
}

fn write_num_cache(writer: &mut Writer<Vec<u8>>, values: &[Option<f64>]) -> Result<()> {
    writer.write_event(Event::Start(BytesStart::new("c:numCache")))?;

    let count_text = values.len().to_string();
    let mut pt_count = BytesStart::new("c:ptCount");
    pt_count.push_attribute(("val", count_text.as_str()));
    writer.write_event(Event::Empty(pt_count))?;

    for (idx, value) in values.iter().enumerate() {
        let Some(value) = value else {
            continue;
        };
        let idx_text = idx.to_string();
        let mut pt = BytesStart::new("c:pt");
        pt.push_attribute(("idx", idx_text.as_str()));
        writer.write_event(Event::Start(pt))?;
        writer.write_event(Event::Start(BytesStart::new("c:v")))?;
        let value_text = value.to_string();
        writer.write_event(Event::Text(BytesText::new(value_text.as_str())))?;
        writer.write_event(Event::End(BytesEnd::new("c:v")))?;
        writer.write_event(Event::End(BytesEnd::new("c:pt")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("c:numCache")))?;
    Ok(())
}

/// Loads charts from drawing relationships for a worksheet.
pub(crate) fn load_worksheet_charts(
    worksheet: &mut Worksheet,
    package: &Package,
    worksheet_uri: &PartUri,
    worksheet_part: &Part,
    drawing_relationship_ids: &[String],
) -> Result<()> {
    for drawing_relationship_id in drawing_relationship_ids {
        let Some(drawing_relationship) = worksheet_part
            .relationships
            .get_by_id(drawing_relationship_id.as_str())
        else {
            continue;
        };

        if drawing_relationship.target_mode != TargetMode::Internal
            || drawing_relationship.rel_type != DRAWING_RELATIONSHIP_TYPE
        {
            continue;
        }

        let drawing_uri = worksheet_uri.resolve_relative(drawing_relationship.target.as_str())?;
        let Some(drawing_part) = package.get_part(drawing_uri.as_str()) else {
            continue;
        };

        for chart_ref in parse_drawing_chart_refs(drawing_part.data.as_bytes())? {
            let Some(chart_relationship) = drawing_part
                .relationships
                .get_by_id(chart_ref.relationship_id.as_str())
            else {
                continue;
            };

            if chart_relationship.target_mode != TargetMode::Internal
                || chart_relationship.rel_type != CHART_RELATIONSHIP_TYPE
            {
                continue;
            }

            let chart_uri = drawing_uri.resolve_relative(chart_relationship.target.as_str())?;
            let Some(chart_part) = package.get_part(chart_uri.as_str()) else {
                continue;
            };

            match parse_chart_xml(chart_part.data.as_bytes()) {
                Ok(mut chart) => {
                    chart.set_from_col(chart_ref.from_col);
                    chart.set_from_row(chart_ref.from_row);
                    chart.set_from_col_off(chart_ref.from_col_off);
                    chart.set_from_row_off(chart_ref.from_row_off);
                    chart.set_to_col(chart_ref.to_col);
                    chart.set_to_row(chart_ref.to_row);
                    chart.set_to_col_off(chart_ref.to_col_off);
                    chart.set_to_row_off(chart_ref.to_row_off);
                    if let Some(ref name) = chart_ref.name {
                        chart.set_name(name.clone());
                    }
                    if let Some(cx) = chart_ref.extent_cx {
                        chart.set_extent_cx(cx);
                    }
                    if let Some(cy) = chart_ref.extent_cy {
                        chart.set_extent_cy(cy);
                    }
                    worksheet.push_chart(chart);
                }
                Err(err) => {
                    tracing::warn!(
                        error = %err,
                        chart_uri = chart_uri.as_str(),
                        "failed to parse chart XML; skipping chart"
                    );
                }
            }
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_chart_xml_bar_chart() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:t>Monthly Sales</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:order val="1"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$C$1</c:f>
            </c:strRef>
          </c:tx>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$C$2:$C$5</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling>
          <c:min val="0"/>
          <c:max val="100"/>
        </c:scaling>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:title>
          <c:tx>
            <c:rich>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p>
                <a:r>
                  <a:t>Revenue ($)</a:t>
                </a:r>
              </a:p>
            </c:rich>
          </c:tx>
        </c:title>
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
</c:chartSpace>"#;

        let chart = parse_chart_xml(xml).expect("should parse chart XML");

        assert_eq!(chart.chart_type(), ChartType::Bar);
        assert_eq!(chart.title(), Some("Monthly Sales"));
        assert_eq!(chart.bar_direction(), Some(BarDirection::Column));
        assert_eq!(chart.grouping(), Some(ChartGrouping::Clustered));
        assert!(!chart.vary_colors());

        // Series
        assert_eq!(chart.series().len(), 2);
        let s0 = &chart.series()[0];
        assert_eq!(s0.idx(), 0);
        assert_eq!(s0.order(), 0);
        assert_eq!(s0.name_ref(), Some("Sheet1!$B$1"));
        assert_eq!(s0.categories().unwrap().formula(), Some("Sheet1!$A$2:$A$5"));
        assert_eq!(s0.values().unwrap().formula(), Some("Sheet1!$B$2:$B$5"));

        let s1 = &chart.series()[1];
        assert_eq!(s1.idx(), 1);
        assert_eq!(s1.order(), 1);
        assert_eq!(s1.name_ref(), Some("Sheet1!$C$1"));
        assert!(s1.categories().is_none());
        assert_eq!(s1.values().unwrap().formula(), Some("Sheet1!$C$2:$C$5"));

        // Axes
        assert_eq!(chart.axes().len(), 2);
        let cat_ax = &chart.axes()[0];
        assert_eq!(cat_ax.id(), 1);
        assert_eq!(cat_ax.axis_type(), "catAx");
        assert_eq!(cat_ax.position(), "b");
        assert_eq!(cat_ax.crosses_ax(), Some(2));
        assert!(!cat_ax.major_gridlines());

        let val_ax = &chart.axes()[1];
        assert_eq!(val_ax.id(), 2);
        assert_eq!(val_ax.axis_type(), "valAx");
        assert_eq!(val_ax.position(), "l");
        assert_eq!(val_ax.min(), Some(0.0));
        assert_eq!(val_ax.max(), Some(100.0));
        assert!(val_ax.major_gridlines());
        assert_eq!(val_ax.title(), Some("Revenue ($)"));
        assert_eq!(val_ax.crosses_ax(), Some(1));

        // Legend
        let legend = chart.legend().expect("should have legend");
        assert_eq!(legend.position(), "r");
        assert!(!legend.overlay());
    }

    #[test]
    fn parse_chart_xml_line_chart_no_legend() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let chart = parse_chart_xml(xml).expect("should parse line chart XML");

        assert_eq!(chart.chart_type(), ChartType::Line);
        assert!(chart.title().is_none());
        assert_eq!(chart.grouping(), Some(ChartGrouping::Standard));
        assert_eq!(chart.series().len(), 1);
        assert!(chart.legend().is_none());
        assert!(chart.axes().is_empty());
    }

    #[test]
    fn parse_chart_xml_pie_chart_3d() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:pie3DChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$1:$A$3</c:f>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$1:$B$3</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:pie3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let chart = parse_chart_xml(xml).expect("should parse pie 3D chart");

        assert_eq!(chart.chart_type(), ChartType::Pie);
        assert!(chart.vary_colors());
    }

    #[test]
    fn parse_chart_xml_preserves_escaped_sheet_names_in_formulas() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:cat>
            <c:strRef>
              <c:f>&apos;Risk Analytics&apos;!$A$18:$A$26</c:f>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>&apos;Risk Analytics&apos;!$D$18:$D$26</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let chart = parse_chart_xml(xml).expect("should parse pie chart");
        let series = &chart.series()[0];

        assert_eq!(
            series.categories().unwrap().formula(),
            Some("'Risk Analytics'!$A$18:$A$26")
        );
        assert_eq!(
            series.values().unwrap().formula(),
            Some("'Risk Analytics'!$D$18:$D$26")
        );
    }

    #[test]
    fn serialize_chart_xml_roundtrip() {
        let mut chart = Chart::new(ChartType::Bar).with_title("Test Chart");
        // Bar charts now get default bar_direction, grouping, and axes

        let mut s = ChartSeries::new(0, 0);
        s.set_name_ref("Sheet1!$B$1")
            .set_categories(ChartDataRef::from_formula("Sheet1!$A$2:$A$5"))
            .set_values(ChartDataRef::from_formula("Sheet1!$B$2:$B$5"));
        chart.add_series(s);

        // Customize the default axes (rather than adding new ones)
        chart.axes_mut()[1]
            .set_title("Amount")
            .set_min(0.0)
            .set_max(1000.0)
            .set_major_gridlines(true);

        let mut legend = ChartLegend::new();
        legend.set_position("r");
        chart.set_legend(legend);

        // Serialize
        let xml_bytes = serialize_chart_xml(&chart).expect("should serialize chart");
        let xml_str = String::from_utf8(xml_bytes.clone()).expect("valid UTF-8");

        // Verify XML contains expected elements
        assert!(xml_str.contains("c:chartSpace"));
        assert!(xml_str.contains("c:barChart"));
        assert!(xml_str.contains("Test Chart"));
        assert!(xml_str.contains("Sheet1!$B$1"));
        assert!(xml_str.contains("Sheet1!$A$2:$A$5"));
        assert!(xml_str.contains("Sheet1!$B$2:$B$5"));
        assert!(xml_str.contains(r#"c:barDir"#));
        assert!(xml_str.contains(r#"val="col"#));
        assert!(xml_str.contains(r#"val="clustered"#));
        assert!(xml_str.contains("c:catAx"));
        assert!(xml_str.contains("c:valAx"));
        assert!(xml_str.contains("Amount"));
        assert!(xml_str.contains("c:majorGridlines"));
        assert!(xml_str.contains("c:legend"));
        assert!(xml_str.contains(r#"c:legendPos"#));

        // Parse back
        let parsed = parse_chart_xml(&xml_bytes).expect("should re-parse chart XML");

        assert_eq!(parsed.chart_type(), ChartType::Bar);
        assert_eq!(parsed.title(), Some("Test Chart"));
        assert_eq!(parsed.bar_direction(), Some(BarDirection::Column));
        assert_eq!(parsed.grouping(), Some(ChartGrouping::Clustered));
        assert_eq!(parsed.series().len(), 1);
        assert_eq!(parsed.series()[0].name_ref(), Some("Sheet1!$B$1"));
        assert_eq!(
            parsed.series()[0].categories().unwrap().formula(),
            Some("Sheet1!$A$2:$A$5")
        );
        assert_eq!(
            parsed.series()[0].values().unwrap().formula(),
            Some("Sheet1!$B$2:$B$5")
        );

        assert_eq!(parsed.axes().len(), 2);
        assert_eq!(parsed.axes()[0].axis_type(), "catAx");
        assert_eq!(parsed.axes()[0].position(), "b");
        assert_eq!(parsed.axes()[1].axis_type(), "valAx");
        assert_eq!(parsed.axes()[1].title(), Some("Amount"));
        assert_eq!(parsed.axes()[1].min(), Some(0.0));
        assert_eq!(parsed.axes()[1].max(), Some(1000.0));
        assert!(parsed.axes()[1].major_gridlines());

        let legend = parsed.legend().expect("should have legend");
        assert_eq!(legend.position(), "r");
    }

    #[test]
    fn serialize_chart_xml_includes_series_caches() {
        let mut chart = Chart::new(ChartType::Bar).with_title("Cached Chart");
        let mut series = ChartSeries::new(0, 0);
        series.set_name("Revenue");

        let mut cats = ChartDataRef::from_formula("Summary!$A$3:$A$6");
        cats.set_str_values(vec![
            "2024-Q1".to_string(),
            "2024-Q2".to_string(),
            "2024-Q3".to_string(),
            "2024-Q4".to_string(),
        ]);

        let mut vals = ChartDataRef::from_formula("Summary!$B$3:$B$6");
        vals.set_num_values(vec![Some(10.0), Some(20.0), Some(30.0), Some(40.0)]);

        series.set_categories(cats).set_values(vals);
        chart.add_series(series);
        chart.add_axis(ChartAxis::new_category());
        chart.add_axis(ChartAxis::new_value());

        let xml_bytes = serialize_chart_xml(&chart).expect("should serialize chart");
        let xml_str = String::from_utf8(xml_bytes).expect("valid UTF-8");

        assert!(xml_str.contains("c:strCache"));
        assert!(xml_str.contains("c:numCache"));
        assert!(xml_str.contains("2024-Q1"));
        assert!(xml_str.contains(">40<"));
        assert!(xml_str.contains("<c:tx>"));
        assert!(xml_str.contains("Revenue"));

        let parsed = parse_chart_xml(xml_str.as_bytes()).expect("should parse chart XML");
        assert_eq!(parsed.series().len(), 1);
        assert_eq!(parsed.series()[0].name(), Some("Revenue"));
        assert_eq!(
            parsed.series()[0].categories().unwrap().str_values(),
            &[
                "2024-Q1".to_string(),
                "2024-Q2".to_string(),
                "2024-Q3".to_string(),
                "2024-Q4".to_string(),
            ]
        );
        assert_eq!(
            parsed.series()[0].values().unwrap().num_values(),
            &[Some(10.0), Some(20.0), Some(30.0), Some(40.0)]
        );
    }

    #[test]
    fn parse_drawing_chart_refs_finds_chart_elements() {
        let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>20</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let refs = parse_drawing_chart_refs(drawing_xml).expect("should parse drawing XML");
        assert_eq!(refs.len(), 2);
        assert_eq!(refs[0].relationship_id, "rId1");
        assert_eq!(refs[1].relationship_id, "rId2");
    }

    #[test]
    fn parse_drawing_chart_refs_empty_when_no_charts() {
        let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="500000" cy="500000"/>
    <xdr:pic>
      <xdr:blipFill><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></xdr:blipFill>
    </xdr:pic>
  </xdr:oneCellAnchor>
</xdr:wsDr>"#;

        let refs = parse_drawing_chart_refs(drawing_xml).expect("should parse drawing XML");
        assert!(refs.is_empty());
    }

    #[test]
    fn parse_drawing_chart_refs_ignores_graphic_frame_ext_for_two_cell_anchors() {
        let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>14</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let refs = parse_drawing_chart_refs(drawing_xml).expect("should parse drawing XML");
        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].relationship_id, "rId1");
        assert_eq!(refs[0].extent_cx, None);
        assert_eq!(refs[0].extent_cy, None);
    }

    #[test]
    fn parse_drawing_chart_refs_reads_outer_ext_for_one_cell_anchors() {
        let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="4572000" cy="2743200"/>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
  </xdr:oneCellAnchor>
</xdr:wsDr>"#;

        let refs = parse_drawing_chart_refs(drawing_xml).expect("should parse drawing XML");
        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].relationship_id, "rId1");
        assert_eq!(refs[0].extent_cx, Some(4_572_000));
        assert_eq!(refs[0].extent_cy, Some(2_743_200));
    }

    #[test]
    fn serialize_chart_xml_with_axis_deleted() {
        let mut chart = Chart::new(ChartType::Line);

        let mut ax = ChartAxis::new_category();
        ax.set_id(1).set_deleted(true);
        chart.add_axis(ax);

        let xml_bytes = serialize_chart_xml(&chart).expect("should serialize");
        let xml_str = String::from_utf8(xml_bytes).expect("valid UTF-8");

        assert!(xml_str.contains(r#"c:delete"#));
        assert!(xml_str.contains(r#"val="1"#));
    }

    #[test]
    fn serialize_chart_xml_with_legend_overlay() {
        let mut chart = Chart::new(ChartType::Pie);
        chart.set_vary_colors(true);

        let mut legend = ChartLegend::new();
        legend.set_position("t").set_overlay(true);
        chart.set_legend(legend);

        let xml_bytes = serialize_chart_xml(&chart).expect("should serialize");
        let xml_str = String::from_utf8(xml_bytes.clone()).expect("valid UTF-8");

        assert!(xml_str.contains("c:pieChart"));
        assert!(xml_str.contains(r#"c:varyColors"#));
        assert!(xml_str.contains(r#"c:overlay"#));

        // Roundtrip
        let parsed = parse_chart_xml(&xml_bytes).expect("should re-parse");
        assert_eq!(parsed.chart_type(), ChartType::Pie);
        assert!(parsed.vary_colors());
        let leg = parsed.legend().expect("legend");
        assert_eq!(leg.position(), "t");
        assert!(leg.overlay());
    }
}
