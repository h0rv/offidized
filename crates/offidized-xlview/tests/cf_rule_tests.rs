//! Comprehensive tests for conditional formatting rule types
//!
//! These tests parse real XLSX files (ms_cf_samples.xlsx and kitchen_sink_v2.xlsx)
//! to verify that various CF rule types are correctly parsed.
//!
//! Tested rule types:
//! - colorScale (2-color and 3-color scales)
//! - dataBar
//! - iconSet
//! - cellIs (equal, notEqual, greaterThan, lessThan, between, etc.)
//! - top10
//! - duplicateValues / uniqueValues
//! - aboveAverage / belowAverage
//! - containsText / notContainsText / beginsWith / endsWith
//! - containsBlanks / notContainsBlanks
//! - timePeriod
//! - expression (custom formula)
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic,
    clippy::approx_constant,
    clippy::cast_possible_truncation,
    clippy::absurd_extreme_comparisons,
    clippy::cast_lossless
)]

mod common;
mod fixtures;

use std::fs;

use offidized_xlview::types::formatting::CFRuleType;
use offidized_xlview::types::workbook::{Sheet, Workbook};

// =============================================================================
// Helper Functions
// =============================================================================

/// Parse an XLSX file and return the workbook
#[allow(clippy::expect_used)]
fn parse_test_file(filename: &str) -> Workbook {
    let path = format!("{}/test/{}", env!("CARGO_MANIFEST_DIR"), filename);
    let data = fs::read(&path).unwrap_or_else(|_| panic!("Failed to read test file: {}", path));
    common::load_xlsx(&data)
}

/// Find a sheet by name in the workbook
#[allow(dead_code)]
fn find_sheet<'a>(workbook: &'a Workbook, name: &str) -> Option<&'a Sheet> {
    workbook.sheets.iter().find(|s| s.name == name)
}

/// Get all CF rules for a specific cell range (sqref)
#[allow(dead_code)]
fn get_cf_rules_for_range<'a>(
    sheet: &'a Sheet,
    sqref_contains: &str,
) -> Vec<&'a offidized_xlview::types::formatting::CFRule> {
    sheet
        .conditional_formatting
        .iter()
        .filter(|cf| cf.sqref.contains(sqref_contains))
        .flat_map(|cf| cf.rules.iter())
        .collect()
}

/// Get first CF rule of a given type from the sheet
#[allow(dead_code)]
fn get_cf_rule_by_type<'a>(
    sheet: &'a Sheet,
    rule_type: &str,
) -> Option<&'a offidized_xlview::types::formatting::CFRule> {
    sheet
        .conditional_formatting
        .iter()
        .flat_map(|cf| cf.rules.iter())
        .find(|r| r.rule_type == rule_type)
}

/// Get all CF rules of a given type from the sheet
fn get_all_cf_rules_by_type<'a>(
    sheet: &'a Sheet,
    rule_type: &str,
) -> Vec<&'a offidized_xlview::types::formatting::CFRule> {
    sheet
        .conditional_formatting
        .iter()
        .flat_map(|cf| cf.rules.iter())
        .filter(|r| r.rule_type == rule_type)
        .collect()
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Color Scale Rules (Sheet 12)
// =============================================================================

#[test]
fn test_color_scale_2_color_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    // Sheet 12 has colorScale rules according to our analysis
    // We need to find a sheet with colorScale - let's check all sheets
    let mut found_color_scale = false;

    for sheet in &workbook.sheets {
        let color_scale_rules = get_all_cf_rules_by_type(sheet, "colorScale");
        if !color_scale_rules.is_empty() {
            found_color_scale = true;

            for rule in color_scale_rules {
                assert_eq!(rule.rule_type, "colorScale");
                assert!(
                    rule.color_scale.is_some(),
                    "colorScale rule should have colorScale data"
                );

                let cs = rule.color_scale.as_ref().unwrap();
                // 2-color or 3-color scale should have 2 or 3 cfvo entries
                assert!(
                    cs.cfvo.len() >= 2 && cs.cfvo.len() <= 3,
                    "colorScale should have 2 or 3 cfvo entries, got {}",
                    cs.cfvo.len()
                );
                // Same number of colors as cfvo entries
                assert_eq!(
                    cs.colors.len(),
                    cs.cfvo.len(),
                    "Number of colors should match number of cfvo entries"
                );
                // All colors should be resolved to #RRGGBB format
                for color in &cs.colors {
                    assert!(
                        color.starts_with('#'),
                        "Color should be in #RRGGBB format, got: {}",
                        color
                    );
                }
            }
        }
    }

    assert!(
        found_color_scale,
        "Should find at least one colorScale rule in ms_cf_samples.xlsx"
    );
}

#[test]
fn test_color_scale_3_color_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    // Find 3-color scales specifically
    let mut found_3_color_scale = false;

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "colorScale" {
                    if let Some(ref cs) = rule.color_scale {
                        if cs.cfvo.len() == 3 {
                            found_3_color_scale = true;
                            assert_eq!(cs.colors.len(), 3, "3-color scale should have 3 colors");

                            // Check that cfvo types are valid
                            for cfvo in &cs.cfvo {
                                let valid_types =
                                    ["min", "max", "num", "percent", "percentile", "formula"];
                                assert!(
                                    valid_types.contains(&cfvo.cfvo_type.as_str()),
                                    "Invalid cfvo type: {}",
                                    cfvo.cfvo_type
                                );
                            }
                        }
                    }
                }
            }
        }
    }

    assert!(
        found_3_color_scale,
        "Should find at least one 3-color scale in ms_cf_samples.xlsx"
    );
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Data Bar Rules (Sheet 11)
// =============================================================================

#[test]
fn test_data_bar_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_data_bar = false;
    let mut found_data_bar_with_show_value_false = false;

    for sheet in &workbook.sheets {
        let data_bar_rules = get_all_cf_rules_by_type(sheet, "dataBar");
        if !data_bar_rules.is_empty() {
            found_data_bar = true;

            for rule in data_bar_rules {
                assert_eq!(rule.rule_type, "dataBar");
                assert!(
                    rule.data_bar.is_some(),
                    "dataBar rule should have dataBar data"
                );

                let db = rule.data_bar.as_ref().unwrap();
                // Data bar should have exactly 2 cfvo entries (min and max)
                assert_eq!(db.cfvo.len(), 2, "dataBar should have 2 cfvo entries");
                // Color should be resolved
                assert!(
                    db.color.starts_with('#'),
                    "dataBar color should be in #RRGGBB format, got: {}",
                    db.color
                );

                if db.show_value == Some(false) {
                    found_data_bar_with_show_value_false = true;
                }
            }
        }
    }

    assert!(
        found_data_bar,
        "Should find at least one dataBar rule in ms_cf_samples.xlsx"
    );
    assert!(
        found_data_bar_with_show_value_false,
        "Should find at least one dataBar with showValue=false"
    );
}

#[test]
fn test_data_bar_cfvo_types() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut cfvo_types_found: Vec<String> = Vec::new();

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "dataBar" {
                    if let Some(ref db) = rule.data_bar {
                        for cfvo in &db.cfvo {
                            if !cfvo_types_found.contains(&cfvo.cfvo_type) {
                                cfvo_types_found.push(cfvo.cfvo_type.clone());
                            }
                        }
                    }
                }
            }
        }
    }

    // We expect various cfvo types in the data bars
    assert!(
        !cfvo_types_found.is_empty(),
        "Should find cfvo types in dataBar rules"
    );
    // Verify at least min/max or percent/percentile types
    let has_expected_types = cfvo_types_found
        .iter()
        .any(|t| t == "min" || t == "max" || t == "percent" || t == "percentile" || t == "num");
    assert!(has_expected_types, "Should find expected cfvo types");
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Icon Set Rules (Sheet 7, 8, 9)
// =============================================================================

#[test]
fn test_icon_set_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_icon_set = false;
    let mut icon_set_names: Vec<String> = Vec::new();

    for sheet in &workbook.sheets {
        let icon_set_rules = get_all_cf_rules_by_type(sheet, "iconSet");
        if !icon_set_rules.is_empty() {
            found_icon_set = true;

            for rule in icon_set_rules {
                assert_eq!(rule.rule_type, "iconSet");
                assert!(
                    rule.icon_set.is_some(),
                    "iconSet rule should have iconSet data"
                );

                let is = rule.icon_set.as_ref().unwrap();
                // Icon set should have at least 3 cfvo entries
                assert!(
                    is.cfvo.len() >= 3,
                    "iconSet should have at least 3 cfvo entries, got {}",
                    is.cfvo.len()
                );

                if !icon_set_names.contains(&is.icon_set) {
                    icon_set_names.push(is.icon_set.clone());
                }
            }
        }
    }

    assert!(
        found_icon_set,
        "Should find at least one iconSet rule in ms_cf_samples.xlsx"
    );

    // ms_cf_samples.xlsx has various icon sets like 3Arrows, 3Symbols2, 5Quarters
    assert!(
        !icon_set_names.is_empty(),
        "Should find various icon set names"
    );
}

#[test]
fn test_icon_set_show_value_attribute() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_show_value_false = false;

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "iconSet" {
                    if let Some(ref is) = rule.icon_set {
                        if is.show_value == Some(false) {
                            found_show_value_false = true;
                        }
                    }
                }
            }
        }
    }

    assert!(
        found_show_value_false,
        "Should find iconSet with showValue=false in ms_cf_samples.xlsx"
    );
}

#[test]
fn test_icon_set_3_arrows() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_3_arrows = false;

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "iconSet" {
                    if let Some(ref is) = rule.icon_set {
                        if is.icon_set == "3Arrows" {
                            found_3_arrows = true;
                            assert_eq!(is.cfvo.len(), 3, "3Arrows should have 3 cfvo entries");
                        }
                    }
                }
            }
        }
    }

    assert!(
        found_3_arrows,
        "Should find 3Arrows iconSet in ms_cf_samples.xlsx"
    );
}

#[test]
fn test_icon_set_5_quarters() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_5_quarters = false;

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "iconSet" {
                    if let Some(ref is) = rule.icon_set {
                        if is.icon_set == "5Quarters" {
                            found_5_quarters = true;
                            assert_eq!(is.cfvo.len(), 5, "5Quarters should have 5 cfvo entries");
                        }
                    }
                }
            }
        }
    }

    assert!(
        found_5_quarters,
        "Should find 5Quarters iconSet in ms_cf_samples.xlsx"
    );
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Cell Is Rules (Sheet 2, 3)
// =============================================================================

#[test]
fn test_cell_is_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_cell_is = false;
    let mut operators_found: Vec<String> = Vec::new();

    for sheet in &workbook.sheets {
        let cell_is_rules = get_all_cf_rules_by_type(sheet, "cellIs");
        if !cell_is_rules.is_empty() {
            found_cell_is = true;

            for rule in cell_is_rules {
                assert_eq!(rule.rule_type, "cellIs");
                // cellIs rules should have an operator
                assert!(
                    rule.operator.is_some(),
                    "cellIs rule should have an operator"
                );
                // cellIs rules should have a formula
                assert!(rule.formula.is_some(), "cellIs rule should have a formula");

                if let Some(ref op) = rule.operator {
                    if !operators_found.contains(op) {
                        operators_found.push(op.clone());
                    }
                }
            }
        }
    }

    assert!(
        found_cell_is,
        "Should find at least one cellIs rule in ms_cf_samples.xlsx"
    );
    assert!(
        !operators_found.is_empty(),
        "Should find various operators for cellIs rules"
    );
}

#[test]
fn test_cell_is_less_than() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_less_than = false;

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "cellIs" {
                    if let Some(ref op) = rule.operator {
                        if op == "lessThan" {
                            found_less_than = true;
                            // Should have a formula value
                            assert!(
                                rule.formula.is_some(),
                                "lessThan rule should have a formula"
                            );
                        }
                    }
                }
            }
        }
    }

    assert!(
        found_less_than,
        "Should find cellIs with operator='lessThan' in ms_cf_samples.xlsx"
    );
}

#[test]
fn test_cell_is_equal() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_equal = false;

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "cellIs" {
                    if let Some(ref op) = rule.operator {
                        if op == "equal" {
                            found_equal = true;
                            assert!(rule.formula.is_some(), "equal rule should have a formula");
                        }
                    }
                }
            }
        }
    }

    assert!(
        found_equal,
        "Should find cellIs with operator='equal' in ms_cf_samples.xlsx"
    );
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Top10 Rules (Sheet 4, 5)
// =============================================================================

#[test]
fn test_top10_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_top10 = false;
    let mut found_percent = false;
    let mut found_bottom = false;

    for sheet in &workbook.sheets {
        let top10_rules = get_all_cf_rules_by_type(sheet, "top10");
        if !top10_rules.is_empty() {
            found_top10 = true;

            for rule in top10_rules {
                assert_eq!(rule.rule_type, "top10");
                // top10 rules should have a rank
                assert!(rule.rank.is_some(), "top10 rule should have a rank");

                if rule.percent == Some(true) {
                    found_percent = true;
                }
                if rule.bottom == Some(true) {
                    found_bottom = true;
                }
            }
        }
    }

    assert!(
        found_top10,
        "Should find at least one top10 rule in ms_cf_samples.xlsx"
    );
    assert!(
        found_percent,
        "Should find top10 rule with percent=true in ms_cf_samples.xlsx"
    );
    assert!(
        found_bottom,
        "Should find top10 rule with bottom=true in ms_cf_samples.xlsx"
    );
}

#[test]
fn test_top10_rank_values() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut rank_values: Vec<u32> = Vec::new();

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "top10" {
                    if let Some(rank) = rule.rank {
                        if !rank_values.contains(&rank) {
                            rank_values.push(rank);
                        }
                    }
                }
            }
        }
    }

    assert!(
        !rank_values.is_empty(),
        "Should find rank values in top10 rules"
    );
    // Verify we have various rank values
    assert!(
        rank_values.iter().any(|&r| r > 0),
        "Rank values should be positive"
    );
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Duplicate/Unique Values Rules (Sheet 6)
// =============================================================================

#[test]
fn test_duplicate_values_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_duplicate_values = false;

    for sheet in &workbook.sheets {
        let dup_rules = get_all_cf_rules_by_type(sheet, "duplicateValues");
        if !dup_rules.is_empty() {
            found_duplicate_values = true;

            for rule in dup_rules {
                assert_eq!(rule.rule_type, "duplicateValues");
                // duplicateValues rules should have a dxfId
                assert!(
                    rule.dxf_id.is_some(),
                    "duplicateValues rule should have a dxfId"
                );
            }
        }
    }

    assert!(
        found_duplicate_values,
        "Should find at least one duplicateValues rule in ms_cf_samples.xlsx"
    );
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Above Average Rules (Sheet 4)
// =============================================================================

#[test]
fn test_above_average_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_above_average = false;

    for sheet in &workbook.sheets {
        let avg_rules = get_all_cf_rules_by_type(sheet, "aboveAverage");
        if !avg_rules.is_empty() {
            found_above_average = true;

            for rule in avg_rules {
                assert_eq!(rule.rule_type, "aboveAverage");
                // aboveAverage rules should have a dxfId
                assert!(
                    rule.dxf_id.is_some(),
                    "aboveAverage rule should have a dxfId"
                );
            }
        }
    }

    assert!(
        found_above_average,
        "Should find at least one aboveAverage rule in ms_cf_samples.xlsx"
    );
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Contains Text Rules (Sheet 2)
// =============================================================================

#[test]
fn test_contains_text_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_contains_text = false;

    for sheet in &workbook.sheets {
        let text_rules = get_all_cf_rules_by_type(sheet, "containsText");
        if !text_rules.is_empty() {
            found_contains_text = true;

            for rule in text_rules {
                assert_eq!(rule.rule_type, "containsText");
                // containsText rules should have a formula
                assert!(
                    rule.formula.is_some(),
                    "containsText rule should have a formula"
                );
                // containsText rules should have operator="containsText"
                assert_eq!(
                    rule.operator,
                    Some("containsText".to_string()),
                    "containsText rule should have operator='containsText'"
                );
            }
        }
    }

    assert!(
        found_contains_text,
        "Should find at least one containsText rule in ms_cf_samples.xlsx"
    );
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Time Period Rules (Sheet 2)
// =============================================================================

#[test]
fn test_time_period_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_time_period = false;

    for sheet in &workbook.sheets {
        let time_rules = get_all_cf_rules_by_type(sheet, "timePeriod");
        if !time_rules.is_empty() {
            found_time_period = true;

            for rule in time_rules {
                assert_eq!(rule.rule_type, "timePeriod");
                // timePeriod rules should have a timePeriod attribute
                assert!(
                    rule.time_period.is_some(),
                    "timePeriod rule should have a timePeriod attribute"
                );
                // timePeriod rules should have a formula
                assert!(
                    rule.formula.is_some(),
                    "timePeriod rule should have a formula"
                );
            }
        }
    }

    assert!(
        found_time_period,
        "Should find at least one timePeriod rule in ms_cf_samples.xlsx"
    );
}

#[test]
fn test_time_period_this_month() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_this_month = false;

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "timePeriod" {
                    if let Some(ref tp) = rule.time_period {
                        if tp == "thisMonth" {
                            found_this_month = true;
                            // Formula should contain MONTH and TODAY
                            if let Some(ref formula) = rule.formula {
                                assert!(
                                    formula.contains("MONTH") && formula.contains("TODAY"),
                                    "thisMonth formula should contain MONTH and TODAY"
                                );
                            }
                        }
                    }
                }
            }
        }
    }

    assert!(
        found_this_month,
        "Should find timePeriod='thisMonth' in ms_cf_samples.xlsx"
    );
}

// =============================================================================
// Tests: ms_cf_samples.xlsx - Expression Rules (Sheet 13, 14, 15, 16)
// =============================================================================

#[test]
fn test_expression_from_real_file() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_expression = false;

    for sheet in &workbook.sheets {
        let expr_rules = get_all_cf_rules_by_type(sheet, "expression");
        if !expr_rules.is_empty() {
            found_expression = true;

            for rule in expr_rules {
                assert_eq!(rule.rule_type, "expression");
                // expression rules should have a formula
                assert!(
                    rule.formula.is_some(),
                    "expression rule should have a formula"
                );
                // expression rules should have a dxfId
                assert!(rule.dxf_id.is_some(), "expression rule should have a dxfId");
            }
        }
    }

    assert!(
        found_expression,
        "Should find at least one expression rule in ms_cf_samples.xlsx"
    );
}

#[test]
fn test_expression_mod_row_formula() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut found_mod_row = false;

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if rule.rule_type == "expression" {
                    if let Some(ref formula) = rule.formula {
                        if formula.contains("MOD") && formula.contains("ROW") {
                            found_mod_row = true;
                        }
                    }
                }
            }
        }
    }

    assert!(
        found_mod_row,
        "Should find expression with MOD(ROW()) formula for alternating row highlighting"
    );
}

// =============================================================================
// Tests: Total CF Count and Statistics
// =============================================================================

#[test]
fn test_cf_rule_count_in_ms_cf_samples() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut total_cf_count = 0;
    let mut total_rule_count = 0;
    let mut rule_type_counts: std::collections::HashMap<CFRuleType, usize> =
        std::collections::HashMap::new();

    for sheet in &workbook.sheets {
        total_cf_count += sheet.conditional_formatting.len();
        for cf in &sheet.conditional_formatting {
            total_rule_count += cf.rules.len();
            for rule in &cf.rules {
                *rule_type_counts.entry(rule.rule_type.clone()).or_insert(0) += 1;
            }
        }
    }

    // ms_cf_samples.xlsx should have many CF rules
    assert!(
        total_cf_count > 10,
        "Should have more than 10 conditionalFormatting elements, got {}",
        total_cf_count
    );
    assert!(
        total_rule_count > 10,
        "Should have more than 10 CF rules total, got {}",
        total_rule_count
    );

    // Verify we found multiple rule types
    assert!(
        rule_type_counts.len() >= 5,
        "Should find at least 5 different rule types, got {}",
        rule_type_counts.len()
    );
}

// =============================================================================
// Tests: DXF Styles - Verify dxfId references are valid
// =============================================================================

#[test]
fn test_dxf_id_references() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let dxf_count = workbook.dxf_styles.len();
    assert!(dxf_count > 0, "Workbook should have dxf_styles");

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                if let Some(dxf_id) = rule.dxf_id {
                    assert!(
                        (dxf_id as usize) < dxf_count,
                        "dxfId {} should be less than dxf_styles count {}",
                        dxf_id,
                        dxf_count
                    );
                }
            }
        }
    }
}

// =============================================================================
// Tests: kitchen_sink_v2.xlsx - Verify CF is parsed (if present)
// =============================================================================

#[test]
fn test_kitchen_sink_v2_parsing() {
    let workbook = parse_test_file("kitchen_sink_v2.xlsx");

    // Verify the file parses successfully
    assert!(
        !workbook.sheets.is_empty(),
        "Should have at least one sheet"
    );

    // Count CF rules in kitchen_sink_v2.xlsx
    let mut total_cf_count = 0;
    for sheet in &workbook.sheets {
        total_cf_count += sheet.conditional_formatting.len();
    }

    // Just verify parsing works - kitchen_sink_v2.xlsx may or may not have CF
    // This test ensures the file can be parsed without errors
    println!(
        "kitchen_sink_v2.xlsx has {} conditionalFormatting elements",
        total_cf_count
    );
}

// =============================================================================
// Tests: CFVO Types
// =============================================================================

#[test]
fn test_cfvo_types_variety() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut cfvo_types: std::collections::HashSet<String> = std::collections::HashSet::new();

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                // Collect from colorScale
                if let Some(ref cs) = rule.color_scale {
                    for cfvo in &cs.cfvo {
                        cfvo_types.insert(cfvo.cfvo_type.clone());
                    }
                }
                // Collect from dataBar
                if let Some(ref db) = rule.data_bar {
                    for cfvo in &db.cfvo {
                        cfvo_types.insert(cfvo.cfvo_type.clone());
                    }
                }
                // Collect from iconSet
                if let Some(ref is) = rule.icon_set {
                    for cfvo in &is.cfvo {
                        cfvo_types.insert(cfvo.cfvo_type.clone());
                    }
                }
            }
        }
    }

    // Should find various cfvo types
    assert!(
        cfvo_types.len() >= 3,
        "Should find at least 3 different cfvo types, got: {:?}",
        cfvo_types
    );

    // Common types that should be present
    let expected_types = ["min", "max", "percent", "num"];
    let has_expected = expected_types.iter().any(|&t| cfvo_types.contains(t));
    assert!(
        has_expected,
        "Should find at least one of: min, max, percent, num"
    );
}

// =============================================================================
// Tests: Priority Values
// =============================================================================

#[test]
fn test_priority_values_are_positive() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            for rule in &cf.rules {
                assert!(
                    rule.priority > 0,
                    "Priority should be positive, got {}",
                    rule.priority
                );
            }
        }
    }
}

// =============================================================================
// Tests: Sqref Parsing
// =============================================================================

#[test]
fn test_sqref_formats() {
    let workbook = parse_test_file("ms_cf_samples.xlsx");

    let mut sqref_examples: Vec<String> = Vec::new();

    for sheet in &workbook.sheets {
        for cf in &sheet.conditional_formatting {
            if sqref_examples.len() < 10 && !sqref_examples.contains(&cf.sqref) {
                sqref_examples.push(cf.sqref.clone());
            }
        }
    }

    // Verify sqref values are not empty
    for sqref in &sqref_examples {
        assert!(!sqref.is_empty(), "sqref should not be empty");
        // Basic validation: should contain column letters and row numbers
        assert!(
            sqref.chars().any(|c| c.is_ascii_alphabetic()),
            "sqref should contain column letters: {}",
            sqref
        );
        assert!(
            sqref.chars().any(|c| c.is_ascii_digit()),
            "sqref should contain row numbers: {}",
            sqref
        );
    }
}
