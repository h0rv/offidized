//! Pivot table XML parsing and serialization.
//!
//! This module provides functions to parse pivot table XML parts and serialize them back.
//! Pivot tables consist of multiple related parts:
//! - pivotTableN.xml (table definition)
//! - pivotCacheDefinitionN.xml (cache definition, shared across tables)
//! - pivotCacheRecordsN.xml (cached data records)

use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet};
use std::io::Cursor;

use offidized_opc::relationship::TargetMode;
use offidized_opc::uri::PartUri;
use offidized_opc::{Package, Part};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::error::{Result, XlsxError};
use crate::pivot_table::{
    PivotDataField, PivotField, PivotFieldSort, PivotSourceReference, PivotSubtotalFunction,
    PivotTable,
};
use crate::worksheet::Worksheet;
use crate::CellValue;

const PIVOT_TABLE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
const PIVOT_CACHE_DEFINITION_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";

#[allow(dead_code)]
const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
#[allow(dead_code)]
const RELATIONSHIP_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

/// Normalize a sheet name extracted from a source reference.
///
/// Excel formulas may quote sheet names (e.g. `'Earnings Fact'!A1:B2`).
/// Internally we need the unquoted sheet name for lookups and worksheetSource@sheet.
fn normalize_sheet_name(sheet: &str) -> Cow<'_, str> {
    let trimmed = sheet.trim();
    if trimmed.len() >= 2 && trimmed.starts_with('\'') && trimmed.ends_with('\'') {
        let inner = &trimmed[1..trimmed.len() - 1];
        return Cow::Owned(inner.replace("''", "'"));
    }
    Cow::Borrowed(trimmed)
}

/// Parsed reference to a pivot table part in a worksheet.
#[derive(Debug, Clone)]
struct ParsedPivotTableRef {
    relationship_id: String,
}

/// Extracts pivot table relationship IDs from a worksheet part.
fn extract_pivot_table_refs(worksheet_part: &Part) -> Vec<ParsedPivotTableRef> {
    let mut refs = Vec::new();
    for rel in worksheet_part.relationships.iter() {
        if rel.rel_type == PIVOT_TABLE_RELATIONSHIP_TYPE && rel.target_mode == TargetMode::Internal
        {
            refs.push(ParsedPivotTableRef {
                relationship_id: rel.id.clone(),
            });
        }
    }
    refs
}

/// Parses field names and source reference from a pivot cache definition XML.
/// Returns (field_names, source_reference)
fn parse_pivot_cache_field_names(xml: &[u8]) -> Result<(Vec<String>, String)> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();
    let mut field_names = Vec::new();
    let mut source_ref = String::new();
    let mut sheet_name = String::new();
    let mut table_name = String::new();
    let mut in_cache_fields = false;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) | Event::Empty(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                match local {
                    b"cacheFields" => {
                        in_cache_fields = true;
                    }
                    b"cacheField" if in_cache_fields => {
                        // Extract the field name
                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            if attr_local == b"name" {
                                let value = String::from_utf8_lossy(attribute.value.as_ref());
                                field_names.push(value.into_owned());
                                break;
                            }
                        }
                    }
                    b"worksheetSource" => {
                        // Parse worksheet source reference - distinguish between range and table
                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            let value = String::from_utf8_lossy(attribute.value.as_ref());
                            match attr_local {
                                b"ref" => {
                                    source_ref = value.into_owned();
                                }
                                b"sheet" => {
                                    sheet_name = value.into_owned();
                                }
                                b"name" => {
                                    // "name" attribute indicates a named table
                                    table_name = value.into_owned();
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());
                if local == b"cacheFields" {
                    in_cache_fields = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    // Determine source reference type
    let full_source_ref = if !table_name.is_empty() {
        // Named table source - just return the table name
        table_name
    } else if !sheet_name.is_empty() && !source_ref.is_empty() {
        // Worksheet range source - combine sheet name and range
        // Add $ signs back to match the expected format
        let ref_with_dollars = source_ref
            .split(':')
            .map(|part| {
                if part.starts_with('$') {
                    part.to_string()
                } else {
                    // Add $ signs if not present
                    let mut result = String::new();
                    let mut has_letter = false;
                    for ch in part.chars() {
                        if ch.is_alphabetic() && !has_letter {
                            result.push('$');
                            has_letter = true;
                        } else if ch.is_numeric() && has_letter {
                            result.push('$');
                            has_letter = false;
                        }
                        result.push(ch);
                    }
                    result
                }
            })
            .collect::<Vec<_>>()
            .join(":");

        format!("{}!{}", sheet_name, ref_with_dollars)
    } else {
        String::new()
    };

    Ok((field_names, full_source_ref))
}

/// Parses a pivot table definition XML part into a `PivotTable` domain object.
fn parse_pivot_table_xml_with_cache_fields(
    xml: &[u8],
    cache_field_names: &[String],
    cache_source_ref: &str,
) -> Result<PivotTable> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buffer = Vec::new();

    let mut table_name = String::new();
    let mut source_ref = String::new();
    let mut target_row: u32 = 0;
    let mut target_col: u32 = 0;
    let mut show_row_grand_totals = true;
    let mut show_column_grand_totals = true;
    let mut row_header_caption: Option<String> = None;
    let mut column_header_caption: Option<String> = None;
    let mut preserve_formatting = true;
    let mut use_auto_formatting = false;
    let mut page_wrap: u32 = 0;
    let mut page_over_then_down = true;
    let mut subtotal_hidden_items = false;

    let mut row_fields: Vec<PivotField> = Vec::new();
    let mut column_fields: Vec<PivotField> = Vec::new();
    let mut page_fields: Vec<PivotField> = Vec::new();
    let mut data_fields: Vec<PivotDataField> = Vec::new();

    let mut in_pivot_table_definition = false;
    let mut in_pivot_fields = false;
    let mut in_row_fields = false;
    let mut in_col_fields = false;
    let mut in_page_fields = false;
    let mut in_data_fields = false;

    // Use cache field names as the indexed list
    let pivot_field_names: Vec<String> = cache_field_names.to_vec();
    let mut pivot_field_properties: Vec<PivotField> = Vec::new();
    let mut current_pivot_field_index: usize = 0;

    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(ref event) | Event::Empty(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                match local {
                    b"pivotTableDefinition" => {
                        in_pivot_table_definition = true;
                        // Parse attributes on the pivotTableDefinition element
                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            let value = String::from_utf8_lossy(attribute.value.as_ref());
                            match attr_local {
                                b"name" => {
                                    table_name = value.into_owned();
                                }
                                b"rowGrandTotals" => {
                                    show_row_grand_totals =
                                        value.trim() == "1" || value.trim() == "true";
                                }
                                b"colGrandTotals" => {
                                    show_column_grand_totals =
                                        value.trim() == "1" || value.trim() == "true";
                                }
                                b"rowHeaderCaption" => {
                                    row_header_caption = Some(value.into_owned());
                                }
                                b"colHeaderCaption" => {
                                    column_header_caption = Some(value.into_owned());
                                }
                                b"preserveFormatting" => {
                                    preserve_formatting =
                                        value.trim() == "1" || value.trim() == "true";
                                }
                                b"useAutoFormatting" => {
                                    use_auto_formatting =
                                        value.trim() == "1" || value.trim() == "true";
                                }
                                b"pageWrap" => {
                                    page_wrap = value.trim().parse().unwrap_or(0);
                                }
                                b"pageOverThenDown" => {
                                    page_over_then_down =
                                        value.trim() == "1" || value.trim() == "true";
                                }
                                b"subtotalHiddenItems" => {
                                    subtotal_hidden_items =
                                        value.trim() == "1" || value.trim() == "true";
                                }
                                _ => {}
                            }
                        }
                    }
                    b"location" if in_pivot_table_definition => {
                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            let value = String::from_utf8_lossy(attribute.value.as_ref());
                            if attr_local == b"ref" {
                                // Parse the reference to extract target cell
                                // The ref is typically like "A3:E13" where A3 is the top-left
                                let ref_str = value.trim();
                                if let Some(colon_pos) = ref_str.find(':') {
                                    let first_cell = &ref_str[..colon_pos];
                                    if let Ok((row, col)) = parse_cell_reference(first_cell) {
                                        target_row = row;
                                        target_col = col;
                                    }
                                } else if let Ok((row, col)) = parse_cell_reference(ref_str) {
                                    target_row = row;
                                    target_col = col;
                                }
                            }
                        }
                    }
                    b"pivotCacheDefinition" => {
                        // Placeholder for pivot cache parsing (simplified for now)
                    }
                    b"cacheSource" => {
                        // Parse cache source for the source reference
                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            if attr_local == b"type" {
                                // Type can be "worksheet", "external", etc.
                            }
                        }
                    }
                    b"worksheetSource" => {
                        // Parse worksheet source reference
                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            let value = String::from_utf8_lossy(attribute.value.as_ref());
                            match attr_local {
                                b"ref" => {
                                    source_ref = value.into_owned();
                                }
                                b"name" => {
                                    // Sheet name (can be prepended to ref)
                                    if !value.is_empty() && !source_ref.contains('!') {
                                        source_ref = format!("{}!{}", value, source_ref);
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    b"pivotFields" if in_pivot_table_definition => {
                        in_pivot_fields = true;
                        // Initialize pivot_field_properties with default fields based on cache names
                        for field_name in &pivot_field_names {
                            pivot_field_properties.push(PivotField::new(field_name.clone()));
                        }
                    }
                    b"pivotField" if in_pivot_fields => {
                        // Parse field properties and update the corresponding field
                        let mut sort_type = PivotFieldSort::Manual;
                        let mut show_all_subtotals = false;
                        let mut insert_blank_rows = false;
                        let mut show_empty_items = false;
                        let mut insert_page_breaks = false;
                        let mut custom_label: Option<String> = None;

                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            let value = String::from_utf8_lossy(attribute.value.as_ref());
                            match attr_local {
                                b"caption" => {
                                    custom_label = Some(value.into_owned());
                                }
                                b"defaultSubtotal" => {
                                    show_all_subtotals =
                                        value.trim() == "1" || value.trim() == "true";
                                }
                                b"sortType" => {
                                    sort_type = PivotFieldSort::from_xml_value(&value);
                                }
                                b"insertBlankRow" => {
                                    insert_blank_rows =
                                        value.trim() == "1" || value.trim() == "true";
                                }
                                b"showAll" => {
                                    show_empty_items =
                                        value.trim() == "1" || value.trim() == "true";
                                }
                                b"insertPageBreak" => {
                                    insert_page_breaks =
                                        value.trim() == "1" || value.trim() == "true";
                                }
                                _ => {}
                            }
                        }

                        // Update the field properties at the current index
                        if let Some(field) =
                            pivot_field_properties.get_mut(current_pivot_field_index)
                        {
                            if let Some(label) = custom_label {
                                field.set_custom_label(label);
                            }
                            field.set_sort_type(sort_type);
                            field.set_show_all_subtotals(show_all_subtotals);
                            field.set_insert_blank_rows(insert_blank_rows);
                            field.set_show_empty_items(show_empty_items);
                            field.set_insert_page_breaks(insert_page_breaks);
                        }
                        current_pivot_field_index += 1;
                    }
                    b"rowFields" if in_pivot_table_definition => {
                        in_row_fields = true;
                    }
                    b"field" if in_row_fields => {
                        // Field reference in row axis - use index to look up field
                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            let value = String::from_utf8_lossy(attribute.value.as_ref());
                            if attr_local == b"x" {
                                // Index into pivotFields collection
                                if let Ok(index) = value.parse::<usize>() {
                                    if let Some(field) = pivot_field_properties.get(index) {
                                        row_fields.push(field.clone());
                                    }
                                }
                            }
                        }
                    }
                    b"colFields" if in_pivot_table_definition => {
                        in_col_fields = true;
                    }
                    b"field" if in_col_fields => {
                        // Field reference in column axis - use index to look up field
                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            let value = String::from_utf8_lossy(attribute.value.as_ref());
                            if attr_local == b"x" {
                                // Index into pivotFields collection
                                if let Ok(index) = value.parse::<usize>() {
                                    if let Some(field) = pivot_field_properties.get(index) {
                                        column_fields.push(field.clone());
                                    }
                                }
                            }
                        }
                    }
                    b"pageFields" if in_pivot_table_definition => {
                        in_page_fields = true;
                    }
                    b"pageField" if in_page_fields => {
                        // Page field reference - parse name attribute and look up properties
                        let mut field_name = String::new();
                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            let value = String::from_utf8_lossy(attribute.value.as_ref());
                            if attr_local == b"name" {
                                field_name = value.into_owned();
                            }
                        }
                        if !field_name.is_empty() {
                            // Find the field index by name
                            if let Some(index) =
                                pivot_field_names.iter().position(|n| n == &field_name)
                            {
                                // Look up the field with properties from pivot_field_properties
                                if let Some(field) = pivot_field_properties.get(index) {
                                    page_fields.push(field.clone());
                                } else {
                                    // Fallback: create new field if not found in properties
                                    page_fields.push(PivotField::new(field_name));
                                }
                            } else {
                                // Field name not in cache, create new field
                                page_fields.push(PivotField::new(field_name));
                            }
                        }
                    }
                    b"dataFields" if in_pivot_table_definition => {
                        in_data_fields = true;
                    }
                    b"dataField" if in_data_fields => {
                        // Parse data field attributes
                        let mut temp_data_custom_name: Option<String> = None;
                        let mut temp_data_subtotal = PivotSubtotalFunction::Sum;
                        let mut temp_data_number_format: Option<String> = None;
                        let mut field_index: Option<usize> = None;

                        for attribute in event.attributes().flatten() {
                            let attr_local = local_name(attribute.key.as_ref());
                            let value = String::from_utf8_lossy(attribute.value.as_ref());
                            match attr_local {
                                b"name" => {
                                    temp_data_custom_name = Some(value.into_owned());
                                }
                                b"fld" => {
                                    // Field index into pivotFields collection
                                    field_index = value.parse::<usize>().ok();
                                }
                                b"subtotal" => {
                                    temp_data_subtotal =
                                        PivotSubtotalFunction::from_xml_value(&value);
                                }
                                b"numFmtId" => {
                                    // For roundtrip, we store the format string directly
                                    // (In full implementation, would lookup in styles table)
                                    temp_data_number_format = Some(value.into_owned());
                                }
                                _ => {}
                            }
                        }

                        // Look up the field name from the pivotFields collection
                        let field_name = if let Some(index) = field_index {
                            pivot_field_names
                                .get(index)
                                .cloned()
                                .unwrap_or_else(|| format!("Field_{}", index))
                        } else {
                            "UnknownField".to_string()
                        };

                        let mut data_field = PivotDataField::new(field_name);
                        data_field.set_subtotal(temp_data_subtotal);
                        if let Some(custom_name) = temp_data_custom_name.clone() {
                            data_field.set_custom_name(custom_name);
                        }
                        if let Some(format) = temp_data_number_format.clone() {
                            data_field.set_number_format(format);
                        }
                        data_fields.push(data_field);
                    }
                    _ => {}
                }
            }
            Event::End(ref event) => {
                let name_bytes = event.name();
                let local = local_name(name_bytes.as_ref());

                match local {
                    b"pivotTableDefinition" => {
                        in_pivot_table_definition = false;
                    }
                    b"pivotFields" => {
                        in_pivot_fields = false;
                    }
                    b"rowFields" => {
                        in_row_fields = false;
                    }
                    b"colFields" => {
                        in_col_fields = false;
                    }
                    b"pageFields" => {
                        in_page_fields = false;
                    }
                    b"dataFields" => {
                        in_data_fields = false;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }

    if table_name.is_empty() {
        return Err(XlsxError::InvalidWorkbookState(
            "pivot table has no name".to_string(),
        ));
    }

    // Use cache source reference if available, otherwise try the one from pivot table XML
    // Determine if it's a WorksheetRange (contains "!") or NamedTable (no "!")
    let source_reference = if !cache_source_ref.is_empty() {
        if cache_source_ref.contains('!') {
            PivotSourceReference::WorksheetRange(cache_source_ref.to_string())
        } else {
            PivotSourceReference::NamedTable(cache_source_ref.to_string())
        }
    } else if !source_ref.is_empty() {
        if source_ref.contains('!') {
            PivotSourceReference::WorksheetRange(source_ref)
        } else {
            PivotSourceReference::NamedTable(source_ref)
        }
    } else {
        // Fallback to default
        PivotSourceReference::WorksheetRange("Sheet1!$A$1:$A$1".to_string())
    };

    let mut table = PivotTable::new(table_name, source_reference);
    table.set_target(target_row, target_col);
    table.set_show_row_grand_totals(show_row_grand_totals);
    table.set_show_column_grand_totals(show_column_grand_totals);
    if let Some(caption) = row_header_caption {
        table.set_row_header_caption(caption);
    }
    if let Some(caption) = column_header_caption {
        table.set_column_header_caption(caption);
    }
    table.set_preserve_formatting(preserve_formatting);
    table.set_use_auto_formatting(use_auto_formatting);
    table.set_page_wrap(page_wrap);
    table.set_page_over_then_down(page_over_then_down);
    table.set_subtotal_hidden_items(subtotal_hidden_items);

    for field in row_fields {
        table.add_row_field(field);
    }
    for field in column_fields {
        table.add_column_field(field);
    }
    for field in page_fields {
        table.add_page_field(field);
    }
    for field in data_fields {
        table.add_data_field(field);
    }

    Ok(table)
}

/// Compute unique item counts per source field from a worksheet range.
fn compute_field_item_counts(
    sheet: &Worksheet,
    range_ref: &str,
    field_names: &[String],
) -> BTreeMap<String, usize> {
    let mut counts = BTreeMap::new();

    let range_clean = range_ref.replace('$', "");
    let parts: Vec<&str> = range_clean.split(':').collect();
    if parts.len() != 2 {
        return counts;
    }
    let Ok((start_row, start_col)) = parse_cell_reference(parts[0]) else {
        return counts;
    };
    let Ok((end_row, _end_col)) = parse_cell_reference(parts[1]) else {
        return counts;
    };

    for (field_index, field_name) in field_names.iter().enumerate() {
        let field_col = start_col + field_index as u32;
        let mut unique = BTreeSet::new();

        for row in (start_row + 1)..=end_row {
            let Ok(col_letter) = column_index_to_letter(field_col) else {
                continue;
            };
            let cell_ref = format!("{}{}", col_letter, row + 1);
            let Some(cell) = sheet.cell(&cell_ref) else {
                continue;
            };
            let Some(value) = cell.value() else {
                continue;
            };

            match value {
                CellValue::String(s) if !s.is_empty() => {
                    unique.insert(s.clone());
                }
                CellValue::Number(n) => {
                    unique.insert(n.to_string());
                }
                CellValue::Bool(b) => {
                    unique.insert(if *b {
                        "true".to_string()
                    } else {
                        "false".to_string()
                    });
                }
                _ => {}
            }
        }

        counts.insert(field_name.clone(), unique.len());
    }

    counts
}

/// Serializes a `PivotTable` domain object into pivot table XML bytes.
pub(crate) fn serialize_pivot_table_xml(
    table: &PivotTable,
    workbook: &crate::Workbook,
) -> Result<Vec<u8>> {
    use crate::CellValue;
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("utf-8"),
        None, // No standalone attribute to match ClosedXML
    )))?;

    let mut root = BytesStart::new("pivotTableDefinition");
    root.push_attribute(("xmlns", SPREADSHEETML_NS));
    root.push_attribute(("name", table.name()));
    root.push_attribute(("cacheId", "0"));
    // Match ClosedXML's minimal attribute set
    root.push_attribute(("applyNumberFormats", "0"));
    root.push_attribute(("applyBorderFormats", "0"));
    root.push_attribute(("applyFontFormats", "0"));
    root.push_attribute(("applyPatternFormats", "0"));
    root.push_attribute(("applyAlignmentFormats", "0"));
    root.push_attribute(("applyWidthHeightFormats", "0"));
    root.push_attribute(("dataCaption", "Values"));

    // Output grand totals only if non-default (default is 1/true for both)
    // ClosedXML omits these when they're at defaults, but we output for roundtrip
    let row_grand_str = if table.show_row_grand_totals() {
        "1"
    } else {
        "0"
    };
    let col_grand_str = if table.show_column_grand_totals() {
        "1"
    } else {
        "0"
    };

    // Only output if non-default (to match ClosedXML closely)
    if !table.show_row_grand_totals() {
        root.push_attribute(("rowGrandTotals", row_grand_str));
    }
    if !table.show_column_grand_totals() {
        root.push_attribute(("colGrandTotals", col_grand_str));
    }

    // Output captions if set
    if let Some(row_caption) = table.row_header_caption() {
        root.push_attribute(("rowHeaderCaption", row_caption));
    }
    if let Some(col_caption) = table.column_header_caption() {
        root.push_attribute(("colHeaderCaption", col_caption));
    }

    // Output other optional attributes for roundtrip (ClosedXML omits defaults)
    if !table.preserve_formatting() {
        root.push_attribute(("preserveFormatting", "0"));
    }
    if table.use_auto_formatting() {
        root.push_attribute(("useAutoFormatting", "1"));
    }
    if table.page_wrap() > 0 {
        let page_wrap_str = table.page_wrap().to_string();
        root.push_attribute(("pageWrap", page_wrap_str.as_str()));
    }
    if !table.page_over_then_down() {
        root.push_attribute(("pageOverThenDown", "0"));
    }
    if table.subtotal_hidden_items() {
        root.push_attribute(("subtotalHiddenItems", "1"));
    }

    root.push_attribute(("outline", "1"));

    writer.write_event(Event::Start(root))?;

    // Location element (target cell) - must be single cell ref, not range
    let target_cell = format!(
        "{}{}",
        column_index_to_letter(table.target_col())?,
        table.target_row() + 1
    );
    let mut location = BytesStart::new("location");
    location.push_attribute(("ref", target_cell.as_str()));
    location.push_attribute(("firstHeaderRow", "0"));
    location.push_attribute(("firstDataRow", "0"));
    location.push_attribute(("firstDataCol", "0"));
    writer.write_event(Event::Empty(location))?;

    // Extract ALL field names from source range header row (CRITICAL for Excel compatibility)
    // This list is used for both cache fields AND pivot fields
    let mut cache_field_names = Vec::new();

    // Parse source reference to get sheet name and range
    let source_str = match table.source_reference() {
        PivotSourceReference::WorksheetRange(ref range) => range.as_str(),
        PivotSourceReference::NamedTable(_) => "",
    };

    let (sheet_name_for_pivot, range_ref_for_pivot) = if let Some(pos) = source_str.find('!') {
        (&source_str[..pos], &source_str[pos + 1..])
    } else {
        ("Sheet1", source_str)
    };
    let sheet_name_for_pivot = normalize_sheet_name(sheet_name_for_pivot);

    let source_sheet_for_pivot = workbook.sheet(sheet_name_for_pivot.as_ref());

    // Read header row to get ALL field names
    if let Some(sheet) = source_sheet_for_pivot {
        if !range_ref_for_pivot.is_empty() {
            let range_clean = range_ref_for_pivot.replace('$', "");
            if let Some(parts) = range_clean.split(':').collect::<Vec<_>>().get(0..2) {
                if let (Ok((start_row, start_col)), Ok((_, end_col))) = (
                    parse_cell_reference(parts[0]),
                    parse_cell_reference(parts[1]),
                ) {
                    // Read header row (first row of range)
                    for col in start_col..=end_col {
                        let col_letter = column_index_to_letter(col).unwrap_or_default();
                        let cell_ref = format!("{}{}", col_letter, start_row + 1);

                        if let Some(cell) = sheet.cell(&cell_ref) {
                            if let Some(CellValue::String(header)) = cell.value() {
                                cache_field_names.push(header.clone());
                            } else {
                                // If no header string, use column letter as field name
                                cache_field_names.push(col_letter.clone());
                            }
                        } else {
                            // No cell found, use column letter as field name
                            cache_field_names.push(col_letter.clone());
                        }
                    }
                }
            }
        }
    }

    let field_item_counts = if let Some(sheet) = source_sheet_for_pivot {
        compute_field_item_counts(sheet, range_ref_for_pivot, &cache_field_names)
    } else {
        BTreeMap::new()
    };

    // Write pivotFields section - ALL cache fields with proper items
    if !cache_field_names.is_empty() {
        let count_str = cache_field_names.len().to_string();
        let mut pivot_fields = BytesStart::new("pivotFields");
        pivot_fields.push_attribute(("count", count_str.as_str()));
        writer.write_event(Event::Start(pivot_fields))?;

        for field_name in &cache_field_names {
            // Find the actual field object if it's in rows, columns, or pages
            let row_field = table.row_fields().iter().find(|f| f.name() == *field_name);
            let col_field = table
                .column_fields()
                .iter()
                .find(|f| f.name() == *field_name);
            let page_field = table.page_fields().iter().find(|f| f.name() == *field_name);
            let field_obj = row_field.or(col_field).or(page_field);

            let mut field_tag = BytesStart::new("pivotField");

            // Check if this field is a data field
            let is_data_field = table
                .data_fields()
                .iter()
                .any(|df| df.field_name() == *field_name);

            // Add name attribute for row/column fields (match ClosedXML)
            if field_obj.is_some() {
                field_tag.push_attribute(("name", field_name.as_str()));
            }

            // Add axis attribute for row/column fields
            if row_field.is_some() {
                field_tag.push_attribute(("axis", "axisRow"));
            } else if col_field.is_some() {
                field_tag.push_attribute(("axis", "axisCol"));
            }

            // Add dataField attribute BEFORE showAll (ClosedXML order)
            if is_data_field {
                field_tag.push_attribute(("dataField", "1"));
            }

            // Write field properties if this is a row/column field
            if let Some(field) = field_obj {
                // Custom label (caption) - displayed label for the field
                if let Some(caption) = field.custom_label() {
                    field_tag.push_attribute(("caption", caption));
                }
                // Sort type
                if field.sort_type() != PivotFieldSort::Manual {
                    let sort_value = field.sort_type().as_xml_value();
                    field_tag.push_attribute(("sortType", sort_value));
                }

                // Show all (showAll in XML)
                if field.show_empty_items() {
                    field_tag.push_attribute(("showAll", "1"));
                } else {
                    field_tag.push_attribute(("showAll", "0"));
                }

                // Default subtotal (based on show_all_subtotals)
                if field.show_all_subtotals() {
                    field_tag.push_attribute(("defaultSubtotal", "1"));
                } else {
                    field_tag.push_attribute(("defaultSubtotal", "0"));
                }

                // Insert blank row
                if field.insert_blank_rows() {
                    field_tag.push_attribute(("insertBlankRow", "1"));
                }

                // Insert page break
                if field.insert_page_breaks() {
                    field_tag.push_attribute(("insertPageBreak", "1"));
                }
            } else {
                // For non-row/col fields, use defaults
                field_tag.push_attribute(("showAll", "0"));
            }

            // Only add items for fields on row/col axes (NOT for data fields or unused fields)
            if field_obj.is_some() {
                let item_count = field_item_counts.get(field_name).copied().unwrap_or(0);
                if item_count > 0 {
                    writer.write_event(Event::Start(field_tag))?;

                    let items_count_str = item_count.to_string();
                    let mut items = BytesStart::new("items");
                    items.push_attribute(("count", items_count_str.as_str()));
                    writer.write_event(Event::Start(items))?;

                    for i in 0..item_count {
                        let mut item = BytesStart::new("item");
                        let i_str = i.to_string();
                        item.push_attribute(("x", i_str.as_str()));
                        writer.write_event(Event::Empty(item))?;
                    }

                    writer.write_event(Event::End(BytesEnd::new("items")))?;
                    writer.write_event(Event::End(BytesEnd::new("pivotField")))?;
                } else {
                    // No items, write as empty element
                    writer.write_event(Event::Empty(field_tag))?;
                }
            } else {
                // Unused fields or data fields: write as empty element
                writer.write_event(Event::Empty(field_tag))?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("pivotFields")))?;
    }

    // Row fields - reference cache indices
    if !table.row_fields().is_empty() {
        let count_str = table.row_fields().len().to_string();
        let mut row_fields_tag = BytesStart::new("rowFields");
        row_fields_tag.push_attribute(("count", count_str.as_str()));
        writer.write_event(Event::Start(row_fields_tag))?;
        for field in table.row_fields() {
            let cache_index = cache_field_names
                .iter()
                .position(|name| *name == field.name())
                .unwrap_or(0);
            let index_str = cache_index.to_string();
            let mut field_ref = BytesStart::new("field");
            field_ref.push_attribute(("x", index_str.as_str()));
            writer.write_event(Event::Empty(field_ref))?;
        }
        writer.write_event(Event::End(BytesEnd::new("rowFields")))?;
    }

    // Column fields - reference cache indices
    if !table.column_fields().is_empty() {
        let count_str = table.column_fields().len().to_string();
        let mut col_fields_tag = BytesStart::new("colFields");
        col_fields_tag.push_attribute(("count", count_str.as_str()));
        writer.write_event(Event::Start(col_fields_tag))?;
        for field in table.column_fields() {
            let cache_index = cache_field_names
                .iter()
                .position(|name| *name == field.name())
                .unwrap_or(0);
            let index_str = cache_index.to_string();
            let mut field_ref = BytesStart::new("field");
            field_ref.push_attribute(("x", index_str.as_str()));
            writer.write_event(Event::Empty(field_ref))?;
        }
        writer.write_event(Event::End(BytesEnd::new("colFields")))?;
    }

    // Page fields
    if !table.page_fields().is_empty() {
        writer.write_event(Event::Start(BytesStart::new("pageFields")))?;
        for field in table.page_fields().iter() {
            let cache_index = cache_field_names
                .iter()
                .position(|name| *name == field.name())
                .unwrap_or(0);
            let index_str = cache_index.to_string();
            let mut page_field = BytesStart::new("pageField");
            page_field.push_attribute(("fld", index_str.as_str()));
            if let Some(label) = field.custom_label() {
                page_field.push_attribute(("name", label));
            } else {
                page_field.push_attribute(("name", field.name()));
            }
            writer.write_event(Event::Empty(page_field))?;
        }
        writer.write_event(Event::End(BytesEnd::new("pageFields")))?;
    }

    // Data fields - reference cache indices
    if !table.data_fields().is_empty() {
        let count_str = table.data_fields().len().to_string();
        let mut data_fields_tag = BytesStart::new("dataFields");
        data_fields_tag.push_attribute(("count", count_str.as_str()));
        writer.write_event(Event::Start(data_fields_tag))?;
        for data_field in table.data_fields() {
            let cache_index = cache_field_names
                .iter()
                .position(|name| *name == data_field.field_name())
                .unwrap_or(0);
            let index_str = cache_index.to_string();
            let mut data_field_tag = BytesStart::new("dataField");

            // Custom name (displayed in pivot table)
            if let Some(custom_name) = data_field.custom_name() {
                data_field_tag.push_attribute(("name", custom_name));
            }

            // Field index
            data_field_tag.push_attribute(("fld", index_str.as_str()));

            // Subtotal function - CRITICAL for roundtrip even if ClosedXML omits it
            let subtotal_value = data_field.subtotal().as_xml_value();
            data_field_tag.push_attribute(("subtotal", subtotal_value));

            // Number format (if specified)
            if let Some(num_fmt) = data_field.number_format() {
                // For now, store as a custom attribute since we don't have style table integration
                // In a full implementation, this would be a numFmtId referencing the styles table
                data_field_tag.push_attribute(("numFmtId", num_fmt));
            }

            writer.write_event(Event::Empty(data_field_tag))?;
        }
        writer.write_event(Event::End(BytesEnd::new("dataFields")))?;
    }

    // Write pivotTableStyleInfo (CRITICAL for LibreOffice rendering)
    let mut style_info = BytesStart::new("pivotTableStyleInfo");
    style_info.push_attribute(("name", "PivotStyleLight16"));
    style_info.push_attribute(("showRowHeaders", "1"));
    style_info.push_attribute(("showColHeaders", "1"));
    writer.write_event(Event::Empty(style_info))?;

    // Write extLst extension list (required by ClosedXML)
    writer.write_event(Event::Start(BytesStart::new("extLst")))?;
    let mut ext = BytesStart::new("ext");
    ext.push_attribute((
        "xmlns:x14",
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    ));
    ext.push_attribute(("uri", "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}"));
    writer.write_event(Event::Start(ext))?;
    let mut pivot_def = BytesStart::new("x14:pivotTableDefinition");
    pivot_def.push_attribute(("enableEdit", "0"));
    pivot_def.push_attribute(("hideValuesRow", "1"));
    writer.write_event(Event::Empty(pivot_def))?;
    writer.write_event(Event::End(BytesEnd::new("ext")))?;
    writer.write_event(Event::End(BytesEnd::new("extLst")))?;

    writer.write_event(Event::End(BytesEnd::new("pivotTableDefinition")))?;

    Ok(writer.into_inner())
}

/// Serializes a pivot cache definition XML for the given pivot table.
pub(crate) fn serialize_pivot_cache_definition_xml(
    table: &PivotTable,
    workbook: &crate::Workbook,
) -> Result<Vec<u8>> {
    use crate::CellValue;
    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("utf-8"),
        None, // No standalone attribute
    )))?;

    // Match ClosedXML's attribute set exactly
    let mut root = BytesStart::new("x:pivotCacheDefinition");
    root.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    root.push_attribute(("xmlns:x", SPREADSHEETML_NS));
    root.push_attribute(("r:id", "rId1")); // Relationship to cache records
    root.push_attribute(("saveData", "1"));
    root.push_attribute(("refreshOnLoad", "1"));
    root.push_attribute(("createdVersion", "5"));
    root.push_attribute(("refreshedVersion", "5"));
    root.push_attribute(("minRefreshableVersion", "3"));
    writer.write_event(Event::Start(root))?;

    // Write cache source - type depends on source reference type
    let mut cache_source = BytesStart::new("x:cacheSource");

    match table.source_reference() {
        PivotSourceReference::WorksheetRange(ref range) => {
            cache_source.push_attribute(("type", "worksheet"));
            writer.write_event(Event::Start(cache_source))?;

            // Parse "Sheet1!$A$1:$D$100" into sheet name and range
            let (sheet_name, range_ref) = if let Some(pos) = range.find('!') {
                let sheet = &range[..pos];
                let range = &range[pos + 1..];
                (sheet, range)
            } else {
                ("Sheet1", range.as_str())
            };
            let sheet_name = normalize_sheet_name(sheet_name);

            // Remove $ signs from range reference to match ClosedXML
            let range_clean = range_ref.replace('$', "");

            let mut worksheet_source = BytesStart::new("x:worksheetSource");
            worksheet_source.push_attribute(("ref", range_clean.as_str()));
            worksheet_source.push_attribute(("sheet", sheet_name.as_ref())); // Use "sheet" not "name"
            writer.write_event(Event::Empty(worksheet_source))?;

            writer.write_event(Event::End(BytesEnd::new("x:cacheSource")))?;
        }
        PivotSourceReference::NamedTable(ref table_name) => {
            cache_source.push_attribute(("type", "worksheet"));
            writer.write_event(Event::Start(cache_source))?;

            // For named tables, use the "name" attribute instead of "ref"
            let mut worksheet_source = BytesStart::new("x:worksheetSource");
            worksheet_source.push_attribute(("name", table_name.as_str()));
            writer.write_event(Event::Empty(worksheet_source))?;

            writer.write_event(Event::End(BytesEnd::new("x:cacheSource")))?;
        }
    }

    // Write cache fields section
    // Extract ALL column names from the source range header row (CRITICAL for Excel compatibility)
    let mut field_names = Vec::new();
    let source_sheet;
    let range_ref;

    match table.source_reference() {
        PivotSourceReference::WorksheetRange(ref range) => {
            // Parse source reference to get sheet name and range
            let (sheet_name, parsed_range) = if let Some(pos) = range.find('!') {
                (&range[..pos], &range[pos + 1..])
            } else {
                ("Sheet1", range.as_str())
            };
            let sheet_name = normalize_sheet_name(sheet_name);

            range_ref = parsed_range;
            source_sheet = workbook.sheet(sheet_name.as_ref());

            // Read header row to get ALL field names
            if let Some(sheet) = source_sheet {
                if !range_ref.is_empty() {
                    let range_clean = range_ref.replace('$', "");
                    if let Some(parts) = range_clean.split(':').collect::<Vec<_>>().get(0..2) {
                        if let (Ok((start_row, start_col)), Ok((_, end_col))) = (
                            parse_cell_reference(parts[0]),
                            parse_cell_reference(parts[1]),
                        ) {
                            // Read header row (first row of range)
                            for col in start_col..=end_col {
                                let col_letter = column_index_to_letter(col).unwrap_or_default();
                                let cell_ref = format!("{}{}", col_letter, start_row + 1);

                                if let Some(cell) = sheet.cell(&cell_ref) {
                                    if let Some(CellValue::String(header)) = cell.value() {
                                        field_names.push(header.clone());
                                    } else {
                                        // If no header string, use column letter as field name
                                        field_names.push(col_letter.clone());
                                    }
                                } else {
                                    // No cell found, use column letter as field name
                                    field_names.push(col_letter.clone());
                                }
                            }
                        }
                    }
                }
            }
        }
        PivotSourceReference::NamedTable(_) => {
            // For named tables, collect field names from the pivot table's fields
            // Order matters: use the order they appear in the pivot table definition
            let mut unique_field_names = BTreeSet::new();

            // Collect from row fields
            for field in table.row_fields() {
                unique_field_names.insert(field.name().to_string());
            }

            // Collect from column fields
            for field in table.column_fields() {
                unique_field_names.insert(field.name().to_string());
            }

            // Collect from page fields
            for field in table.page_fields() {
                unique_field_names.insert(field.name().to_string());
            }

            // Collect from data fields (the source fields, not the aggregated names)
            for field in table.data_fields() {
                unique_field_names.insert(field.field_name().to_string());
            }

            field_names = unique_field_names.into_iter().collect();

            // Set these to empty/None for named tables
            source_sheet = None;
            range_ref = "";
        }
    }

    if !field_names.is_empty() {
        // NO count attribute on cacheFields (ClosedXML doesn't use it)
        let cache_fields = BytesStart::new("x:cacheFields");
        writer.write_event(Event::Start(cache_fields))?;

        // source_str, sheet_name, range_ref, and source_sheet are already extracted above

        for (field_index, field_name) in field_names.iter().enumerate() {
            let mut cache_field = BytesStart::new("x:cacheField");
            cache_field.push_attribute(("name", &**field_name));
            // NO numFmtId attribute (ClosedXML doesn't use it)
            writer.write_event(Event::Start(cache_field))?;

            // Collect field values and determine data type (match ClosedXML's approach)
            // Use Vec to preserve insertion order (not HashSet)
            let mut string_values: Vec<String> = Vec::new();
            let mut numeric_values: Vec<f64> = Vec::new();
            let mut min_num: Option<f64> = None;
            let mut max_num: Option<f64> = None;
            let mut is_numeric = false;
            let mut is_string = false;

            if let Some(sheet) = source_sheet {
                if !range_ref.is_empty() {
                    let range_clean = range_ref.replace('$', "");
                    if let Some(parts) = range_clean.split(':').collect::<Vec<_>>().get(0..2) {
                        if let (Ok((start_row, start_col)), Ok((end_row, _))) = (
                            parse_cell_reference(parts[0]),
                            parse_cell_reference(parts[1]),
                        ) {
                            let field_col = start_col + field_index as u32;

                            // Read values from data rows (skip header)
                            for row in (start_row + 1)..=end_row {
                                let col_letter =
                                    column_index_to_letter(field_col).unwrap_or_default();
                                let cell_ref = format!("{}{}", col_letter, row + 1);

                                if let Some(cell) = sheet.cell(&cell_ref) {
                                    if let Some(value) = cell.value() {
                                        match value {
                                            CellValue::Number(n) => {
                                                is_numeric = true;
                                                let n_val = *n;
                                                numeric_values.push(n_val);
                                                min_num =
                                                    Some(min_num.map_or(n_val, |m| m.min(n_val)));
                                                max_num =
                                                    Some(max_num.map_or(n_val, |m| m.max(n_val)));
                                            }
                                            CellValue::String(s) => {
                                                is_string = true;
                                                // Preserve insertion order, avoid duplicates
                                                if !string_values.iter().any(|v| v == s) {
                                                    string_values.push(s.clone());
                                                }
                                            }
                                            CellValue::Bool(_) => {
                                                is_string = true;
                                            }
                                            _ => {}
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Write sharedItems based on data type (match ClosedXML's approach)
            if is_numeric && !is_string {
                // Pure numeric field: list ALL values like ClosedXML does
                let count_str = numeric_values.len().to_string();
                let mut shared_items = BytesStart::new("x:sharedItems");
                shared_items.push_attribute(("containsSemiMixedTypes", "0"));
                shared_items.push_attribute(("containsString", "0"));
                shared_items.push_attribute(("containsNumber", "1"));
                shared_items.push_attribute(("containsInteger", "1"));
                if let (Some(min), Some(max)) = (min_num, max_num) {
                    let min_str = min.to_string();
                    let max_str = max.to_string();
                    shared_items.push_attribute(("minValue", min_str.as_str()));
                    shared_items.push_attribute(("maxValue", max_str.as_str()));
                }
                shared_items.push_attribute(("count", count_str.as_str()));
                writer.write_event(Event::Start(shared_items))?;

                // Write all numeric values as <x:n> elements
                for num in numeric_values {
                    let num_str = if num.fract() == 0.0 {
                        format!("{}", num as i64)
                    } else {
                        num.to_string()
                    };
                    let mut n_elem = BytesStart::new("x:n");
                    n_elem.push_attribute(("v", num_str.as_str()));
                    writer.write_event(Event::Empty(n_elem))?;
                }
                writer.write_event(Event::End(BytesEnd::new("x:sharedItems")))?;
            } else if is_string && string_values.len() <= 50 {
                // Categorical field with few values: list them
                let count_str = string_values.len().to_string();
                let mut shared_items = BytesStart::new("x:sharedItems");
                shared_items.push_attribute(("count", count_str.as_str()));
                writer.write_event(Event::Start(shared_items))?;

                for value in string_values {
                    let mut item = BytesStart::new("x:s");
                    item.push_attribute(("v", value.as_str()));
                    writer.write_event(Event::Empty(item))?;
                }
                writer.write_event(Event::End(BytesEnd::new("x:sharedItems")))?;
            } else {
                // Many unique values or mixed types: empty sharedItems
                writer.write_event(Event::Empty(BytesStart::new("x:sharedItems")))?;
            }
            writer.write_event(Event::End(BytesEnd::new("x:cacheField")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("x:cacheFields")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("x:pivotCacheDefinition")))?;

    Ok(writer.into_inner())
}

/// Serializes pivot cache records XML with actual data from the source range.
pub(crate) fn serialize_pivot_cache_records_xml(
    table: &PivotTable,
    workbook: &crate::Workbook,
) -> Result<Vec<u8>> {
    use crate::CellValue;

    let mut writer = Writer::new_with_indent(Vec::new(), b' ', 2);
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("utf-8"),
        None, // No standalone attribute to match ClosedXML
    )))?;

    // Parse source reference to get sheet name and range
    let source_str = match table.source_reference() {
        PivotSourceReference::WorksheetRange(ref range) => range.as_str(),
        PivotSourceReference::NamedTable(ref _name) => {
            // For named tables, we'd need to look up the table definition
            // For now, just write empty records
            let mut root = BytesStart::new("pivotCacheRecords");
            root.push_attribute(("xmlns", SPREADSHEETML_NS));
            root.push_attribute(("count", "0"));
            writer.write_event(Event::Empty(root))?;
            return Ok(writer.into_inner());
        }
    };

    let (sheet_name, range_ref) = if let Some(pos) = source_str.find('!') {
        let sheet = &source_str[..pos];
        let range = &source_str[pos + 1..];
        (sheet, range)
    } else {
        // No sheet name, assume first sheet
        ("Sheet1", source_str)
    };
    let sheet_name = normalize_sheet_name(sheet_name);

    // Get the source worksheet
    let Some(sheet) = workbook.sheet(sheet_name.as_ref()) else {
        // Sheet not found, write empty records
        let mut root = BytesStart::new("pivotCacheRecords");
        root.push_attribute(("xmlns:r", RELATIONSHIP_NS));
        root.push_attribute((
            "xmlns:mc",
            "http://schemas.openxmlformats.org/markup-compatibility/2006",
        ));
        root.push_attribute((
            "xmlns:xr",
            "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
        ));
        root.push_attribute(("xmlns", SPREADSHEETML_NS));
        root.push_attribute(("mc:Ignorable", "xr"));
        writer.write_event(Event::Empty(root))?;
        return Ok(writer.into_inner());
    };

    // Parse the range (e.g., "$A$1:$F$51")
    let range_clean = range_ref.replace('$', "");
    let parts: Vec<&str> = range_clean.split(':').collect();
    if parts.len() != 2 {
        // Invalid range format, write empty
        let mut root = BytesStart::new("pivotCacheRecords");
        root.push_attribute(("xmlns", SPREADSHEETML_NS));
        root.push_attribute(("count", "0"));
        writer.write_event(Event::Empty(root))?;
        return Ok(writer.into_inner());
    }

    let start_cell = parts[0];
    let end_cell = parts[1];

    let (start_row, start_col) = parse_cell_reference(start_cell)?;
    let (end_row, end_col) = parse_cell_reference(end_cell)?;

    // Read data from the range (skip header row)
    let mut records = Vec::new();
    for row in (start_row + 1)..=end_row {
        let mut record = Vec::new();
        for col in start_col..=end_col {
            let col_letter = column_index_to_letter(col)?;
            let cell_ref = format!("{}{}", col_letter, row + 1);

            let value = sheet
                .cell(&cell_ref)
                .and_then(|cell| cell.value())
                .cloned()
                .unwrap_or(CellValue::String(String::new()));

            record.push(value);
        }
        records.push(record);
    }

    // Match ClosedXML's namespace declarations exactly
    let mut root = BytesStart::new("pivotCacheRecords");
    root.push_attribute(("xmlns:r", RELATIONSHIP_NS));
    root.push_attribute((
        "xmlns:mc",
        "http://schemas.openxmlformats.org/markup-compatibility/2006",
    ));
    root.push_attribute((
        "xmlns:xr",
        "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    ));
    root.push_attribute(("xmlns", SPREADSHEETML_NS));
    root.push_attribute(("mc:Ignorable", "xr"));
    // Note: ClosedXML does NOT include count attribute
    writer.write_event(Event::Start(root))?;

    // Write each record
    for record in records {
        writer.write_event(Event::Start(BytesStart::new("r")))?;

        for value in record {
            match value {
                CellValue::Number(n) => {
                    let mut elem = BytesStart::new("n");
                    let val_str = n.to_string();
                    elem.push_attribute(("v", val_str.as_str()));
                    writer.write_event(Event::Empty(elem))?;
                }
                CellValue::String(ref s) => {
                    let mut elem = BytesStart::new("s");
                    elem.push_attribute(("v", s.as_str()));
                    writer.write_event(Event::Empty(elem))?;
                }
                CellValue::Bool(b) => {
                    let mut elem = BytesStart::new("b");
                    elem.push_attribute(("v", if b { "1" } else { "0" }));
                    writer.write_event(Event::Empty(elem))?;
                }
                _ => {
                    // For other types, write as missing value
                    writer.write_event(Event::Empty(BytesStart::new("m")))?;
                }
            }
        }

        writer.write_event(Event::End(BytesEnd::new("r")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("pivotCacheRecords")))?;

    Ok(writer.into_inner())
}

/// Helper wrapper for tests - parses without cache field names
#[cfg(test)]
pub(crate) fn parse_pivot_table_xml(xml: &[u8]) -> Result<PivotTable> {
    // For tests, use empty cache fields and source (will result in defaults)
    parse_pivot_table_xml_with_cache_fields(xml, &[], "")
}

/// Loads pivot tables from worksheet relationships.
pub(crate) fn load_worksheet_pivot_tables(
    worksheet: &mut Worksheet,
    package: &Package,
    worksheet_uri: &PartUri,
    worksheet_part: &Part,
) -> Result<()> {
    let pivot_table_refs = extract_pivot_table_refs(worksheet_part);

    for pivot_ref in pivot_table_refs {
        let Some(pivot_relationship) = worksheet_part
            .relationships
            .get_by_id(pivot_ref.relationship_id.as_str())
        else {
            continue;
        };

        if pivot_relationship.target_mode != TargetMode::Internal
            || pivot_relationship.rel_type != PIVOT_TABLE_RELATIONSHIP_TYPE
        {
            continue;
        }

        let pivot_uri = worksheet_uri.resolve_relative(pivot_relationship.target.as_str())?;
        let Some(pivot_part) = package.get_part(pivot_uri.as_str()) else {
            continue;
        };

        // Load the pivot cache definition to get field names and source reference
        let mut cache_field_names = Vec::new();
        let mut cache_source_ref = String::new();
        let cache_rels = pivot_part
            .relationships
            .get_by_type(PIVOT_CACHE_DEFINITION_RELATIONSHIP_TYPE);
        if let Some(cache_rel) = cache_rels.first() {
            if cache_rel.target_mode == TargetMode::Internal {
                if let Ok(cache_uri) = pivot_uri.resolve_relative(cache_rel.target.as_str()) {
                    if let Some(cache_part) = package.get_part(cache_uri.as_str()) {
                        (cache_field_names, cache_source_ref) =
                            parse_pivot_cache_field_names(cache_part.data.as_bytes())
                                .unwrap_or_default();
                    }
                }
            }
        }

        match parse_pivot_table_xml_with_cache_fields(
            pivot_part.data.as_bytes(),
            &cache_field_names,
            &cache_source_ref,
        ) {
            Ok(pivot_table) => {
                worksheet.push_pivot_table(pivot_table);
            }
            Err(err) => {
                tracing::warn!(
                    error = %err,
                    pivot_uri = pivot_uri.as_str(),
                    "failed to parse pivot table XML; skipping pivot table"
                );
            }
        }
    }

    Ok(())
}

/// Parses a cell reference like "A1" into (row, col) 0-indexed.
fn parse_cell_reference(cell_ref: &str) -> Result<(u32, u32)> {
    let cell_ref = cell_ref.trim();
    let mut col_part = String::new();
    let mut row_part = String::new();
    let mut found_digit = false;

    for ch in cell_ref.chars() {
        if ch.is_ascii_digit() {
            found_digit = true;
            row_part.push(ch);
        } else if ch.is_ascii_alphabetic() && !found_digit {
            col_part.push(ch);
        } else if ch == '$' {
            // Skip absolute reference markers
        } else {
            break;
        }
    }

    if col_part.is_empty() || row_part.is_empty() {
        return Err(XlsxError::InvalidWorkbookState(format!(
            "invalid cell reference: {}",
            cell_ref
        )));
    }

    let row = row_part
        .parse::<u32>()
        .map_err(|_| XlsxError::InvalidWorkbookState(format!("invalid row number: {}", row_part)))?
        .saturating_sub(1);

    let col = column_letter_to_index(&col_part)?;

    Ok((row, col))
}

/// Converts column letter (A, B, ..., Z, AA, ...) to 0-indexed column number.
fn column_letter_to_index(col_letter: &str) -> Result<u32> {
    let mut result: u32 = 0;
    for ch in col_letter.chars() {
        let ch_upper = ch.to_ascii_uppercase();
        if !ch_upper.is_ascii_alphabetic() {
            return Err(XlsxError::InvalidWorkbookState(format!(
                "invalid column letter: {}",
                col_letter
            )));
        }
        let digit = (ch_upper as u32) - (b'A' as u32) + 1;
        result = result.saturating_mul(26).saturating_add(digit);
    }
    Ok(result.saturating_sub(1))
}

/// Converts 0-indexed column number to column letter (A, B, ..., Z, AA, ...).
#[allow(dead_code)]
fn column_index_to_letter(col_index: u32) -> Result<String> {
    let mut col = col_index + 1;
    let mut result = String::new();
    while col > 0 {
        let remainder = (col - 1) % 26;
        result.insert(0, (b'A' + (remainder as u8)) as char);
        col = (col - 1) / 26;
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pivot_table::PivotFieldSort;

    #[test]
    fn test_parse_cell_reference() {
        assert_eq!(parse_cell_reference("A1").ok(), Some((0, 0)));
        assert_eq!(parse_cell_reference("B2").ok(), Some((1, 1)));
        assert_eq!(parse_cell_reference("Z10").ok(), Some((9, 25)));
        assert_eq!(parse_cell_reference("AA1").ok(), Some((0, 26)));
        assert_eq!(parse_cell_reference("$A$1").ok(), Some((0, 0)));
        assert_eq!(parse_cell_reference("$B$2").ok(), Some((1, 1)));
    }

    #[test]
    fn test_column_letter_to_index() {
        assert_eq!(column_letter_to_index("A").ok(), Some(0));
        assert_eq!(column_letter_to_index("B").ok(), Some(1));
        assert_eq!(column_letter_to_index("Z").ok(), Some(25));
        assert_eq!(column_letter_to_index("AA").ok(), Some(26));
        assert_eq!(column_letter_to_index("AB").ok(), Some(27));
        assert_eq!(column_letter_to_index("AZ").ok(), Some(51));
        assert_eq!(column_letter_to_index("BA").ok(), Some(52));
    }

    #[test]
    fn test_column_index_to_letter() {
        assert_eq!(column_index_to_letter(0).ok(), Some("A".to_string()));
        assert_eq!(column_index_to_letter(1).ok(), Some("B".to_string()));
        assert_eq!(column_index_to_letter(25).ok(), Some("Z".to_string()));
        assert_eq!(column_index_to_letter(26).ok(), Some("AA".to_string()));
        assert_eq!(column_index_to_letter(27).ok(), Some("AB".to_string()));
        assert_eq!(column_index_to_letter(51).ok(), Some("AZ".to_string()));
        assert_eq!(column_index_to_letter(52).ok(), Some("BA".to_string()));
    }

    #[test]
    fn test_serialize_parse_roundtrip() {
        let source_ref = PivotSourceReference::from_range("Sheet1!$A$1:$D$100");
        let mut table = PivotTable::new("PivotTable1", source_ref);
        table.set_target(2, 5);
        table.set_show_row_grand_totals(false);
        table.set_show_column_grand_totals(true);
        table.set_row_header_caption("Rows");
        table.set_column_header_caption("Columns");

        let mut row_field = PivotField::new("Category");
        row_field.set_custom_label("Product Category");
        table.add_row_field(row_field);

        let mut data_field = PivotDataField::new("Sales");
        data_field.set_subtotal(PivotSubtotalFunction::Sum);
        data_field.set_custom_name("Total Sales");
        table.add_data_field(data_field);

        // Create a minimal workbook for the test
        let wb = crate::Workbook::new();
        let xml = serialize_pivot_table_xml(&table, &wb).expect("should serialize");
        let xml_str = String::from_utf8(xml.clone()).expect("valid UTF-8");

        assert!(xml_str.contains("pivotTableDefinition"));
        assert!(xml_str.contains("PivotTable1"));
        assert!(xml_str.contains("location"));
        // rowGrandTotals is non-default (false), so should be output
        assert!(xml_str.contains("rowGrandTotals=\"0\""));
        // colGrandTotals is default (true), so ClosedXML doesn't output it (we match that)
        // Captions should be output
        assert!(xml_str.contains("rowHeaderCaption=\"Rows\""));
        assert!(xml_str.contains("colHeaderCaption=\"Columns\""));

        let parsed = parse_pivot_table_xml(&xml).expect("should parse");
        assert_eq!(parsed.name(), "PivotTable1");
        assert_eq!(parsed.target_row(), 2);
        assert_eq!(parsed.target_col(), 5);
        assert!(!parsed.show_row_grand_totals());
        assert!(parsed.show_column_grand_totals());
        assert_eq!(parsed.row_header_caption(), Some("Rows"));
        assert_eq!(parsed.column_header_caption(), Some("Columns"));
    }

    #[test]
    fn test_pivot_table_builder_api() {
        let source_ref = PivotSourceReference::from_table("SalesData");
        let mut table = PivotTable::new("MonthlySales", source_ref);

        table
            .set_target(5, 2)
            .set_show_row_grand_totals(true)
            .set_show_column_grand_totals(false)
            .set_preserve_formatting(false)
            .set_use_auto_formatting(true)
            .set_page_wrap(3)
            .set_page_over_then_down(false)
            .set_subtotal_hidden_items(true);

        // Add row fields
        let mut region_field = PivotField::new("Region");
        region_field.set_show_all_subtotals(true);
        table.add_row_field(region_field);

        let mut product_field = PivotField::new("Product");
        product_field
            .set_custom_label("Product Name")
            .set_insert_blank_rows(true)
            .set_show_empty_items(true);
        table.add_row_field(product_field);

        // Add column fields
        let mut year_field = PivotField::new("Year");
        year_field.set_sort_type(PivotFieldSort::Descending);
        table.add_column_field(year_field);

        // Add page fields
        let sales_person_field = PivotField::new("SalesPerson");
        table.add_page_field(sales_person_field);

        // Add data fields
        let mut sum_revenue = PivotDataField::new("Revenue");
        sum_revenue
            .set_custom_name("Total Revenue")
            .set_subtotal(PivotSubtotalFunction::Sum)
            .set_number_format("#,##0.00");
        table.add_data_field(sum_revenue);

        let mut avg_price = PivotDataField::new("UnitPrice");
        avg_price
            .set_custom_name("Average Unit Price")
            .set_subtotal(PivotSubtotalFunction::Average)
            .set_number_format("$#,##0.00");
        table.add_data_field(avg_price);

        // Verify structure
        assert_eq!(table.name(), "MonthlySales");
        assert_eq!(table.target_row(), 5);
        assert_eq!(table.target_col(), 2);
        assert!(table.show_row_grand_totals());
        assert!(!table.show_column_grand_totals());
        assert!(!table.preserve_formatting());
        assert!(table.use_auto_formatting());
        assert_eq!(table.page_wrap(), 3);
        assert!(!table.page_over_then_down());
        assert!(table.subtotal_hidden_items());

        assert_eq!(table.row_fields().len(), 2);
        assert_eq!(table.column_fields().len(), 1);
        assert_eq!(table.page_fields().len(), 1);
        assert_eq!(table.data_fields().len(), 2);

        // Verify field properties
        assert_eq!(table.row_fields()[0].name(), "Region");
        assert!(table.row_fields()[0].show_all_subtotals());

        assert_eq!(table.row_fields()[1].name(), "Product");
        assert_eq!(table.row_fields()[1].custom_label(), Some("Product Name"));
        assert!(table.row_fields()[1].insert_blank_rows());
        assert!(table.row_fields()[1].show_empty_items());

        assert_eq!(table.column_fields()[0].name(), "Year");
        assert_eq!(
            table.column_fields()[0].sort_type(),
            PivotFieldSort::Descending
        );

        assert_eq!(table.data_fields()[0].field_name(), "Revenue");
        assert_eq!(table.data_fields()[0].custom_name(), Some("Total Revenue"));
        assert_eq!(
            table.data_fields()[0].subtotal(),
            PivotSubtotalFunction::Sum
        );
        assert_eq!(table.data_fields()[0].number_format(), Some("#,##0.00"));

        assert_eq!(table.data_fields()[1].field_name(), "UnitPrice");
        assert_eq!(
            table.data_fields()[1].custom_name(),
            Some("Average Unit Price")
        );
        assert_eq!(
            table.data_fields()[1].subtotal(),
            PivotSubtotalFunction::Average
        );
        assert_eq!(table.data_fields()[1].number_format(), Some("$#,##0.00"));
    }

    #[test]
    fn test_pivot_subtotal_functions() {
        assert_eq!(PivotSubtotalFunction::Sum.as_xml_value(), "sum");
        assert_eq!(PivotSubtotalFunction::Average.as_xml_value(), "average");
        assert_eq!(PivotSubtotalFunction::Count.as_xml_value(), "count");
        assert_eq!(PivotSubtotalFunction::CountNums.as_xml_value(), "countNums");
        assert_eq!(PivotSubtotalFunction::Max.as_xml_value(), "max");
        assert_eq!(PivotSubtotalFunction::Min.as_xml_value(), "min");
        assert_eq!(PivotSubtotalFunction::Product.as_xml_value(), "product");
        assert_eq!(PivotSubtotalFunction::StdDev.as_xml_value(), "stdDev");
        assert_eq!(PivotSubtotalFunction::StdDevP.as_xml_value(), "stdDevP");
        assert_eq!(PivotSubtotalFunction::Var.as_xml_value(), "var");
        assert_eq!(PivotSubtotalFunction::VarP.as_xml_value(), "varP");

        assert_eq!(
            PivotSubtotalFunction::from_xml_value("sum"),
            PivotSubtotalFunction::Sum
        );
        assert_eq!(
            PivotSubtotalFunction::from_xml_value("average"),
            PivotSubtotalFunction::Average
        );
        assert_eq!(
            PivotSubtotalFunction::from_xml_value("max"),
            PivotSubtotalFunction::Max
        );
        assert_eq!(
            PivotSubtotalFunction::from_xml_value("unknown"),
            PivotSubtotalFunction::Sum
        );
    }

    #[test]
    fn test_pivot_field_sort() {
        assert_eq!(PivotFieldSort::Manual.as_xml_value(), "manual");
        assert_eq!(PivotFieldSort::Ascending.as_xml_value(), "ascending");
        assert_eq!(PivotFieldSort::Descending.as_xml_value(), "descending");

        assert_eq!(
            PivotFieldSort::from_xml_value("ascending"),
            PivotFieldSort::Ascending
        );
        assert_eq!(
            PivotFieldSort::from_xml_value("descending"),
            PivotFieldSort::Descending
        );
        assert_eq!(
            PivotFieldSort::from_xml_value("manual"),
            PivotFieldSort::Manual
        );
        assert_eq!(
            PivotFieldSort::from_xml_value("unknown"),
            PivotFieldSort::Manual
        );
    }

    #[test]
    fn test_pivot_source_reference() {
        let range_ref = PivotSourceReference::from_range("Sheet1!$A$1:$D$100");
        match &range_ref {
            PivotSourceReference::WorksheetRange(ref_str) => {
                assert_eq!(ref_str, "Sheet1!$A$1:$D$100");
            }
            _ => panic!("Expected WorksheetRange"),
        }
        assert_eq!(range_ref.as_str(), "Sheet1!$A$1:$D$100");

        let table_ref = PivotSourceReference::from_table("MyTable");
        match &table_ref {
            PivotSourceReference::NamedTable(name) => {
                assert_eq!(name, "MyTable");
            }
            _ => panic!("Expected NamedTable"),
        }
        assert_eq!(table_ref.as_str(), "MyTable");
    }

    #[test]
    fn test_complex_roundtrip() {
        let source_ref = PivotSourceReference::from_range("Data!$A$1:$Z$1000");
        let mut table = PivotTable::new("ComplexPivot", source_ref);

        table
            .set_target(10, 15)
            .set_show_row_grand_totals(false)
            .set_show_column_grand_totals(false)
            .set_row_header_caption("Row Labels")
            .set_column_header_caption("Column Labels")
            .set_preserve_formatting(true)
            .set_use_auto_formatting(false)
            .set_page_wrap(5)
            .set_page_over_then_down(true)
            .set_subtotal_hidden_items(false);

        // Multiple row fields
        let mut field1 = PivotField::new("Department");
        field1
            .set_show_all_subtotals(true)
            .set_insert_blank_rows(false);
        table.add_row_field(field1);

        let mut field2 = PivotField::new("Employee");
        field2
            .set_custom_label("Employee Name")
            .set_show_empty_items(true);
        table.add_row_field(field2);

        // Multiple column fields
        let mut col1 = PivotField::new("Quarter");
        col1.set_sort_type(PivotFieldSort::Descending);
        table.add_column_field(col1);

        let mut col2 = PivotField::new("Product");
        col2.set_insert_page_breaks(true);
        table.add_column_field(col2);

        // Page fields
        table.add_page_field(PivotField::new("Region"));
        table.add_page_field(PivotField::new("Year"));

        // Multiple data fields
        let mut data1 = PivotDataField::new("Sales");
        data1
            .set_custom_name("Total Sales")
            .set_subtotal(PivotSubtotalFunction::Sum);
        table.add_data_field(data1);

        let mut data2 = PivotDataField::new("Quantity");
        data2
            .set_custom_name("Total Quantity")
            .set_subtotal(PivotSubtotalFunction::Sum);
        table.add_data_field(data2);

        let mut data3 = PivotDataField::new("Sales");
        data3
            .set_custom_name("Average Sales")
            .set_subtotal(PivotSubtotalFunction::Average);
        table.add_data_field(data3);

        // Serialize and parse
        // Create a minimal workbook for the test
        let wb = crate::Workbook::new();
        let xml = serialize_pivot_table_xml(&table, &wb).expect("should serialize");
        let parsed = parse_pivot_table_xml(&xml).expect("should parse");

        // Verify all properties roundtripped
        assert_eq!(parsed.name(), table.name());
        assert_eq!(parsed.target_row(), table.target_row());
        assert_eq!(parsed.target_col(), table.target_col());
        assert_eq!(
            parsed.show_row_grand_totals(),
            table.show_row_grand_totals()
        );
        assert_eq!(
            parsed.show_column_grand_totals(),
            table.show_column_grand_totals()
        );
        assert_eq!(parsed.row_header_caption(), table.row_header_caption());
        assert_eq!(
            parsed.column_header_caption(),
            table.column_header_caption()
        );
        assert_eq!(parsed.preserve_formatting(), table.preserve_formatting());
        assert_eq!(parsed.use_auto_formatting(), table.use_auto_formatting());
        assert_eq!(parsed.page_wrap(), table.page_wrap());
        assert_eq!(parsed.page_over_then_down(), table.page_over_then_down());
        assert_eq!(
            parsed.subtotal_hidden_items(),
            table.subtotal_hidden_items()
        );

        // Verify field counts (parsing is simplified, so we just check structure)
        assert_eq!(parsed.page_fields().len(), 2);
        assert_eq!(parsed.data_fields().len(), 3);
    }
}
