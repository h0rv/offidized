//! Regression tests for bug fixes and UX improvements.

use offidized_xlsx::{range::CellRange, CellValue, Workbook};

/// Test that set_formula() automatically strips leading '=' for better UX.
/// Users often include the '=' prefix, and we should handle both cases.
#[test]
fn test_formula_auto_strip_equals() {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_sheet("Sheet1");

    // Test with leading '='
    sheet
        .cell_mut("A1")
        .expect("cell should exist")
        .set_formula("=SUM(1,2,3)");

    // Test without leading '='
    sheet
        .cell_mut("A2")
        .expect("cell should exist")
        .set_formula("SUM(4,5,6)");

    // Both should have formulas WITHOUT the leading '='
    let cell_a1 = sheet.cell("A1").expect("cell should exist");
    assert_eq!(cell_a1.formula(), Some("SUM(1,2,3)"));

    let cell_a2 = sheet.cell("A2").expect("cell should exist");
    assert_eq!(cell_a2.formula(), Some("SUM(4,5,6)"));

    // Verify roundtrip preserves formulas correctly
    let temp_dir = tempfile::tempdir().expect("tempdir should be created");
    let path = temp_dir.path().join("formula_strip_test.xlsx");
    workbook.save(&path).expect("save should succeed");

    let loaded = Workbook::open(&path).expect("open should succeed");
    let loaded_sheet = loaded.sheet("Sheet1").expect("sheet should exist");

    let loaded_a1 = loaded_sheet.cell("A1").expect("cell should exist");
    assert_eq!(loaded_a1.formula(), Some("SUM(1,2,3)"));

    let loaded_a2 = loaded_sheet.cell("A2").expect("cell should exist");
    assert_eq!(loaded_a2.formula(), Some("SUM(4,5,6)"));
}

/// Test that range copy works correctly with values, formulas, and styles.
#[test]
fn test_range_copy_regression() {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_sheet("Sheet1");

    // Set up source data
    sheet.cell_mut("A1").expect("cell").set_value("Header");
    sheet.cell_mut("A2").expect("cell").set_value("Item 1");
    sheet.cell_mut("A3").expect("cell").set_value("Item 2");
    sheet.cell_mut("B2").expect("cell").set_value(100);
    sheet.cell_mut("B3").expect("cell").set_value(200);
    sheet.cell_mut("C2").expect("cell").set_formula("B2*2");
    sheet.cell_mut("C3").expect("cell").set_formula("B3*2");

    // Copy range A1:C3 to E1
    let range = CellRange::parse("A1:C3").expect("range should parse");
    range.copy_to(sheet, "E1").expect("copy should succeed");

    // Verify copied values
    assert_eq!(
        sheet.cell("E1").expect("cell").value(),
        Some(&CellValue::String("Header".to_string()))
    );
    assert_eq!(
        sheet.cell("E2").expect("cell").value(),
        Some(&CellValue::String("Item 1".to_string()))
    );
    assert_eq!(
        sheet.cell("E3").expect("cell").value(),
        Some(&CellValue::String("Item 2".to_string()))
    );
    assert_eq!(
        sheet.cell("F2").expect("cell").value(),
        Some(&CellValue::Number(100.0))
    );
    assert_eq!(
        sheet.cell("F3").expect("cell").value(),
        Some(&CellValue::Number(200.0))
    );

    // Verify copied formulas are adjusted correctly (implemented in range copy feature)
    assert_eq!(sheet.cell("G2").expect("cell").formula(), Some("F2*2"));
    assert_eq!(sheet.cell("G3").expect("cell").formula(), Some("F3*2"));

    // Verify roundtrip
    let temp_dir = tempfile::tempdir().expect("tempdir should be created");
    let path = temp_dir.path().join("range_copy_test.xlsx");
    workbook.save(&path).expect("save should succeed");

    let loaded = Workbook::open(&path).expect("open should succeed");
    let loaded_sheet = loaded.sheet("Sheet1").expect("sheet should exist");

    assert_eq!(
        loaded_sheet.cell("E1").expect("cell").value(),
        Some(&CellValue::String("Header".to_string()))
    );
    assert_eq!(
        loaded_sheet.cell("G2").expect("cell").formula(),
        Some("F2*2")
    );
}

/// Test that pivot tables are written with cache definition and cache records.
#[test]
fn test_pivot_table_cache_serialization() {
    use offidized_xlsx::{
        PivotDataField, PivotField, PivotSourceReference, PivotSubtotalFunction, PivotTable,
    };

    let mut workbook = Workbook::new();

    // Add source data
    let sales = workbook.add_sheet("Sheet1");
    sales.cell_mut("A1").expect("cell").set_value("Region");
    sales.cell_mut("B1").expect("cell").set_value("Product");
    sales.cell_mut("C1").expect("cell").set_value("Revenue");
    sales.cell_mut("A2").expect("cell").set_value("North");
    sales.cell_mut("B2").expect("cell").set_value("Widget");
    sales.cell_mut("C2").expect("cell").set_value(1000);

    // Add pivot table
    let pivot_sheet = workbook.add_sheet("Pivot");
    let source = PivotSourceReference::from_range("Sheet1!$A$1:$C$2");
    let mut pivot = PivotTable::new("TestPivot", source);
    pivot.set_target(0, 0);
    pivot.add_row_field(PivotField::new("Region"));

    let mut revenue_field = PivotDataField::new("Revenue");
    revenue_field.set_subtotal(PivotSubtotalFunction::Sum);
    pivot.add_data_field(revenue_field);

    pivot_sheet.add_pivot_table(pivot);

    // Save and verify pivot cache parts exist
    let temp_dir = tempfile::tempdir().expect("tempdir should be created");
    let path = temp_dir.path().join("pivot_cache_test.xlsx");
    workbook.save(&path).expect("save should succeed");

    // Verify file contains pivot cache parts
    let file = std::fs::File::open(&path).expect("file should open");
    let mut archive = zip::ZipArchive::new(file).expect("zip should open");

    // Debug: List all files in ZIP to verify cache parts are saved
    eprintln!("\n=== Files in saved workbook ZIP ===");
    for i in 0..archive.len() {
        let file = archive.by_index(i).unwrap();
        eprintln!("  {}", file.name());
    }
    eprintln!("=== End of file list ===\n");

    // Check pivot table part exists (first table is "pivotTable.xml" per ClosedXML naming)
    assert!(
        archive.by_name("xl/pivotTables/pivotTable.xml").is_ok(),
        "pivot table part should exist"
    );

    // Check pivot cache definition exists
    assert!(
        archive
            .by_name("xl/pivotCache/pivotCacheDefinition1.xml")
            .is_ok(),
        "pivot cache definition should exist"
    );

    // Check pivot cache records exists
    assert!(
        archive
            .by_name("xl/pivotCache/pivotCacheRecords1.xml")
            .is_ok(),
        "pivot cache records should exist"
    );

    // Verify cache definition contains source reference
    let mut cache_def = archive
        .by_name("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("cache def should exist");
    let mut cache_xml = String::new();
    std::io::Read::read_to_string(&mut cache_def, &mut cache_xml).expect("read should succeed");

    assert!(
        cache_xml.contains("Sheet1"),
        "cache definition should contain source sheet name"
    );
    assert!(
        cache_xml.contains("A1:C2"),
        "cache definition should contain source range"
    );
}

/// Test that quoted sheet names in pivot source ranges still resolve cache data.
#[test]
fn test_pivot_table_cache_with_quoted_sheet_name() {
    use offidized_xlsx::{
        PivotDataField, PivotField, PivotSourceReference, PivotSubtotalFunction, PivotTable,
    };

    let mut workbook = Workbook::new();
    let sales = workbook.add_sheet("Earnings Fact");
    sales.cell_mut("A1").expect("cell").set_value("Region");
    sales.cell_mut("B1").expect("cell").set_value("Revenue");
    sales
        .cell_mut("A2")
        .expect("cell")
        .set_value("North America");
    sales.cell_mut("B2").expect("cell").set_value(1234);

    let pivot_sheet = workbook.add_sheet("Pivot");
    let source = PivotSourceReference::from_range("'Earnings Fact'!$A$1:$B$2");
    let mut pivot = PivotTable::new("QuotedSheetPivot", source);
    pivot.set_target(0, 0);
    pivot.add_row_field(PivotField::new("Region"));
    let mut revenue_field = PivotDataField::new("Revenue");
    revenue_field.set_subtotal(PivotSubtotalFunction::Sum);
    pivot.add_data_field(revenue_field);
    pivot_sheet.add_pivot_table(pivot);

    let temp_dir = tempfile::tempdir().expect("tempdir should be created");
    let path = temp_dir.path().join("pivot_quoted_sheet_test.xlsx");
    workbook.save(&path).expect("save should succeed");

    let file = std::fs::File::open(&path).expect("file should open");
    let mut archive = zip::ZipArchive::new(file).expect("zip should open");

    let mut cache_def = archive
        .by_name("xl/pivotCache/pivotCacheDefinition1.xml")
        .expect("cache definition should exist");
    let mut cache_def_xml = String::new();
    std::io::Read::read_to_string(&mut cache_def, &mut cache_def_xml).expect("read should work");
    drop(cache_def);

    assert!(
        cache_def_xml.contains("sheet=\"Earnings Fact\""),
        "worksheetSource should use an unquoted sheet name"
    );

    let mut cache_records = archive
        .by_name("xl/pivotCache/pivotCacheRecords1.xml")
        .expect("cache records should exist");
    let mut cache_records_xml = String::new();
    std::io::Read::read_to_string(&mut cache_records, &mut cache_records_xml)
        .expect("read should work");

    assert!(
        cache_records_xml.contains("North America"),
        "cache records should include source data rows"
    );
}

/// Test that pageField indexes map to cache field indexes (not page-field order).
#[test]
fn test_pivot_page_field_index_matches_source_field_index() {
    use offidized_xlsx::{
        PivotDataField, PivotField, PivotSourceReference, PivotSubtotalFunction, PivotTable,
    };

    let mut workbook = Workbook::new();
    let data = workbook.add_sheet("Data");
    data.cell_mut("A1").expect("cell").set_value("Region");
    data.cell_mut("B1").expect("cell").set_value("Product");
    data.cell_mut("C1").expect("cell").set_value("Quarter");
    data.cell_mut("D1").expect("cell").set_value("Revenue");
    data.cell_mut("A2").expect("cell").set_value("North");
    data.cell_mut("B2").expect("cell").set_value("Platform");
    data.cell_mut("C2").expect("cell").set_value("Q1");
    data.cell_mut("D2").expect("cell").set_value(10);

    let pivot_sheet = workbook.add_sheet("Pivot");
    let source = PivotSourceReference::from_range("Data!$A$1:$D$2");
    let mut pivot = PivotTable::new("PivotWithPageField", source);
    pivot.set_target(0, 0);
    pivot.add_row_field(PivotField::new("Region"));
    pivot.add_column_field(PivotField::new("Quarter"));
    pivot.add_page_field(PivotField::new("Product")); // Source index 1
    let mut revenue_field = PivotDataField::new("Revenue");
    revenue_field.set_subtotal(PivotSubtotalFunction::Sum);
    pivot.add_data_field(revenue_field);
    pivot_sheet.add_pivot_table(pivot);

    let temp_dir = tempfile::tempdir().expect("tempdir should be created");
    let path = temp_dir.path().join("pivot_page_field_index_test.xlsx");
    workbook.save(&path).expect("save should succeed");

    let pivot_xml = {
        let file = std::fs::File::open(&path).expect("file should open");
        let mut archive = zip::ZipArchive::new(file).expect("zip should open");
        let mut part = archive
            .by_name("xl/pivotTables/pivotTable.xml")
            .expect("pivot part should exist");
        let mut xml = String::new();
        std::io::Read::read_to_string(&mut part, &mut xml).expect("read should succeed");
        xml
    };

    assert!(
        pivot_xml.contains(r#"<pageField fld="1" name="Product"/>"#),
        "pageField fld should reference Product source-field index (1)"
    );
}

/// Test that conditional formatting works and roundtrips correctly.
#[test]
fn test_conditional_formatting_regression() {
    use offidized_xlsx::{ConditionalFormatting, ConditionalFormattingOperator};

    let mut workbook = Workbook::new();
    let sheet = workbook.add_sheet("Sheet1");

    // Add values
    for row in 2..=11 {
        sheet
            .cell_mut(&format!("A{}", row))
            .expect("cell")
            .set_value((row - 1) * 10);
    }

    // Add conditional formatting (highlight > 50)
    let mut cf =
        ConditionalFormatting::cell_is(vec!["A2:A11"], vec!["50"]).expect("cf should be created");
    cf.set_operator(ConditionalFormattingOperator::GreaterThan);
    sheet.add_conditional_formatting(cf);

    // Verify roundtrip
    let temp_dir = tempfile::tempdir().expect("tempdir should be created");
    let path = temp_dir.path().join("conditional_formatting_test.xlsx");
    workbook.save(&path).expect("save should succeed");

    let loaded = Workbook::open(&path).expect("open should succeed");
    let loaded_sheet = loaded.sheet("Sheet1").expect("sheet should exist");

    // Verify conditional formatting was preserved
    assert_eq!(
        loaded_sheet.conditional_formattings().len(),
        1,
        "conditional formatting should be preserved"
    );
}
