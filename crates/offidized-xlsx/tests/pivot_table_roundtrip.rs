#![allow(clippy::unwrap_used)]
#![allow(clippy::expect_used)]
#![allow(clippy::bool_assert_comparison)]

use offidized_xlsx::{
    CellValue, PivotDataField, PivotField, PivotFieldSort, PivotSourceReference,
    PivotSubtotalFunction, PivotTable, Workbook,
};
use tempfile::TempDir;

/// Helper function to create sample sales data in a worksheet.
fn create_sample_sales_data(workbook: &mut Workbook, sheet_name: &str) {
    let sheet = workbook.add_sheet(sheet_name);

    // Headers
    sheet.cell_mut("A1").unwrap().set_value("Region");
    sheet.cell_mut("B1").unwrap().set_value("Product");
    sheet.cell_mut("C1").unwrap().set_value("Category");
    sheet.cell_mut("D1").unwrap().set_value("Sales");
    sheet.cell_mut("E1").unwrap().set_value("Units");
    sheet.cell_mut("F1").unwrap().set_value("Price");

    // Data rows
    let data = [
        ("North", "Widget", "Electronics", 1000, 10, 100),
        ("North", "Gadget", "Electronics", 1500, 15, 100),
        ("South", "Widget", "Electronics", 2000, 20, 100),
        ("South", "Gadget", "Electronics", 2500, 25, 100),
        ("East", "Widget", "Home", 1200, 12, 100),
        ("East", "Tool", "Home", 1800, 18, 100),
        ("West", "Gadget", "Electronics", 3000, 30, 100),
        ("West", "Tool", "Home", 900, 9, 100),
    ];

    for (idx, (region, product, category, sales, units, price)) in data.iter().enumerate() {
        let row = idx + 2;
        sheet
            .cell_mut(&format!("A{}", row))
            .unwrap()
            .set_value(*region);
        sheet
            .cell_mut(&format!("B{}", row))
            .unwrap()
            .set_value(*product);
        sheet
            .cell_mut(&format!("C{}", row))
            .unwrap()
            .set_value(*category);
        sheet
            .cell_mut(&format!("D{}", row))
            .unwrap()
            .set_value(*sales);
        sheet
            .cell_mut(&format!("E{}", row))
            .unwrap()
            .set_value(*units);
        sheet
            .cell_mut(&format!("F{}", row))
            .unwrap()
            .set_value(*price);
    }
}

#[test]
fn test_pivot_table_creation_and_access() {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_sheet("Data");

    // Add some sample data
    sheet.cell_mut("A1").unwrap().set_value("Region");
    sheet.cell_mut("B1").unwrap().set_value("Product");
    sheet.cell_mut("C1").unwrap().set_value("Sales");
    sheet.cell_mut("A2").unwrap().set_value("North");
    sheet.cell_mut("B2").unwrap().set_value("Widget");
    sheet.cell_mut("C2").unwrap().set_value(1000);
    sheet.cell_mut("A3").unwrap().set_value("South");
    sheet.cell_mut("B3").unwrap().set_value("Gadget");
    sheet.cell_mut("C3").unwrap().set_value(1500);

    // Create a pivot table
    let source_ref = PivotSourceReference::from_range("Data!$A$1:$C$3");
    let mut pivot_table = PivotTable::new("SalesPivot", source_ref);
    pivot_table.set_target(5, 0);

    // Add row field
    let mut region_field = PivotField::new("Region");
    region_field.set_show_all_subtotals(true);
    pivot_table.add_row_field(region_field);

    // Add data field
    let mut sales_data = PivotDataField::new("Sales");
    sales_data
        .set_custom_name("Total Sales")
        .set_subtotal(PivotSubtotalFunction::Sum);
    pivot_table.add_data_field(sales_data);

    // Add pivot table to sheet
    sheet.add_pivot_table(pivot_table);

    // Verify access
    assert_eq!(sheet.pivot_tables().len(), 1);
    assert_eq!(sheet.pivot_tables()[0].name(), "SalesPivot");
    assert_eq!(sheet.pivot_tables()[0].target_row(), 5);
    assert_eq!(sheet.pivot_tables()[0].target_col(), 0);
    assert_eq!(sheet.pivot_tables()[0].row_fields().len(), 1);
    assert_eq!(sheet.pivot_tables()[0].data_fields().len(), 1);
}

#[test]
fn test_pivot_table_roundtrip_in_workbook() {
    let temp_dir = TempDir::new().expect("failed to create temp dir");
    let file_path = temp_dir.path().join("pivot_test.xlsx");

    // Create workbook with pivot table
    {
        let mut workbook = Workbook::new();
        let sheet = workbook.add_sheet("Sales");

        // Add data
        sheet.cell_mut("A1").unwrap().set_value("Department");
        sheet.cell_mut("B1").unwrap().set_value("Employee");
        sheet.cell_mut("C1").unwrap().set_value("Revenue");
        sheet.cell_mut("A2").unwrap().set_value("Engineering");
        sheet.cell_mut("B2").unwrap().set_value("Alice");
        sheet.cell_mut("C2").unwrap().set_value(75000);
        sheet.cell_mut("A3").unwrap().set_value("Engineering");
        sheet.cell_mut("B3").unwrap().set_value("Bob");
        sheet.cell_mut("C3").unwrap().set_value(80000);
        sheet.cell_mut("A4").unwrap().set_value("Sales");
        sheet.cell_mut("B4").unwrap().set_value("Charlie");
        sheet.cell_mut("C4").unwrap().set_value(90000);

        // Create pivot table
        let source_ref = PivotSourceReference::from_range("Sales!$A$1:$C$4");
        let mut pivot_table = PivotTable::new("RevenuePivot", source_ref);
        pivot_table
            .set_target(6, 0)
            .set_show_row_grand_totals(true)
            .set_show_column_grand_totals(false)
            .set_row_header_caption("Departments");

        // Add row fields
        let mut dept_field = PivotField::new("Department");
        dept_field.set_show_all_subtotals(true);
        pivot_table.add_row_field(dept_field);

        let mut emp_field = PivotField::new("Employee");
        emp_field
            .set_custom_label("Employee Name")
            .set_show_empty_items(true);
        pivot_table.add_row_field(emp_field);

        // Add data field
        let mut revenue_data = PivotDataField::new("Revenue");
        revenue_data
            .set_custom_name("Total Revenue")
            .set_subtotal(PivotSubtotalFunction::Sum)
            .set_number_format("#,##0");
        pivot_table.add_data_field(revenue_data);

        sheet.add_pivot_table(pivot_table);

        workbook.save(&file_path).expect("failed to save workbook");
    }

    // Verify file exists
    assert!(file_path.exists(), "workbook file should exist");

    // Reload and verify
    let workbook = Workbook::open(&file_path).expect("failed to open workbook");
    let sheet = workbook.sheet("Sales").expect("sheet should exist");

    // Verify pivot table was preserved
    assert_eq!(sheet.pivot_tables().len(), 1);
    let pivot = &sheet.pivot_tables()[0];
    assert_eq!(pivot.name(), "RevenuePivot");
    assert_eq!(pivot.target_row(), 6);
    assert_eq!(pivot.target_col(), 0);
    assert_eq!(pivot.show_row_grand_totals(), true);
    assert_eq!(pivot.show_column_grand_totals(), false);
    assert_eq!(pivot.row_header_caption(), Some("Departments"));
    assert_eq!(pivot.row_fields().len(), 2);
    assert_eq!(pivot.data_fields().len(), 1);

    // Verify row fields
    assert_eq!(pivot.row_fields()[0].name(), "Department");
    assert_eq!(pivot.row_fields()[0].show_all_subtotals(), true);
    assert_eq!(pivot.row_fields()[1].name(), "Employee");
    assert_eq!(pivot.row_fields()[1].custom_label(), Some("Employee Name"));
    assert_eq!(pivot.row_fields()[1].show_empty_items(), true);

    // Verify data field
    assert_eq!(pivot.data_fields()[0].field_name(), "Revenue");
    assert_eq!(pivot.data_fields()[0].custom_name(), Some("Total Revenue"));
    assert_eq!(
        pivot.data_fields()[0].subtotal(),
        PivotSubtotalFunction::Sum
    );
    assert_eq!(pivot.data_fields()[0].number_format(), Some("#,##0"));
}

#[test]
fn test_multiple_pivot_tables_on_same_worksheet() {
    let temp_dir = TempDir::new().expect("failed to create temp dir");
    let file_path = temp_dir.path().join("multi_pivot_test.xlsx");

    // Create workbook with multiple pivot tables on same sheet
    {
        let mut workbook = Workbook::new();
        create_sample_sales_data(&mut workbook, "Data");
        let sheet = workbook.sheet_mut("Data").unwrap();

        // Create first pivot table - Sales by Region
        let source1 = PivotSourceReference::from_range("Data!$A$1:$F$9");
        let mut pivot1 = PivotTable::new("SalesByRegion", source1);
        pivot1.set_target(12, 0);
        let mut region_field = PivotField::new("Region");
        region_field.set_sort_type(PivotFieldSort::Ascending);
        pivot1.add_row_field(region_field);
        let mut sales_data = PivotDataField::new("Sales");
        sales_data
            .set_custom_name("Total Sales")
            .set_subtotal(PivotSubtotalFunction::Sum);
        pivot1.add_data_field(sales_data);
        sheet.add_pivot_table(pivot1);

        // Create second pivot table - Units by Product
        let source2 = PivotSourceReference::from_range("Data!$A$1:$F$9");
        let mut pivot2 = PivotTable::new("UnitsByProduct", source2);
        pivot2.set_target(12, 5);
        let mut product_field = PivotField::new("Product");
        product_field.set_sort_type(PivotFieldSort::Descending);
        pivot2.add_row_field(product_field);
        let mut units_data = PivotDataField::new("Units");
        units_data
            .set_custom_name("Total Units")
            .set_subtotal(PivotSubtotalFunction::Sum);
        pivot2.add_data_field(units_data);
        sheet.add_pivot_table(pivot2);

        // Create third pivot table - Category breakdown
        let source3 = PivotSourceReference::from_range("Data!$A$1:$F$9");
        let mut pivot3 = PivotTable::new("CategoryBreakdown", source3);
        pivot3.set_target(12, 10);
        let category_field = PivotField::new("Category");
        pivot3.add_row_field(category_field);
        let mut avg_price_data = PivotDataField::new("Price");
        avg_price_data
            .set_custom_name("Avg Price")
            .set_subtotal(PivotSubtotalFunction::Average);
        pivot3.add_data_field(avg_price_data);
        sheet.add_pivot_table(pivot3);

        // Create fourth pivot table - Complex multi-field
        let source4 = PivotSourceReference::from_range("Data!$A$1:$F$9");
        let mut pivot4 = PivotTable::new("ComplexPivot", source4);
        pivot4.set_target(25, 0);
        let region_field2 = PivotField::new("Region");
        pivot4.add_row_field(region_field2);
        let category_field2 = PivotField::new("Category");
        pivot4.add_row_field(category_field2);
        let mut sales_sum = PivotDataField::new("Sales");
        sales_sum
            .set_custom_name("Sum")
            .set_subtotal(PivotSubtotalFunction::Sum);
        pivot4.add_data_field(sales_sum);
        sheet.add_pivot_table(pivot4);

        workbook.save(&file_path).expect("failed to save workbook");
    }

    // Reload and verify all pivot tables preserved
    let workbook = Workbook::open(&file_path).expect("failed to open workbook");
    let sheet = workbook.sheet("Data").expect("sheet should exist");

    assert_eq!(sheet.pivot_tables().len(), 4);

    // Verify first pivot table
    let pivot1 = &sheet.pivot_tables()[0];
    assert_eq!(pivot1.name(), "SalesByRegion");
    assert_eq!(pivot1.target_row(), 12);
    assert_eq!(pivot1.target_col(), 0);
    assert_eq!(pivot1.row_fields().len(), 1);
    assert_eq!(pivot1.row_fields()[0].name(), "Region");
    assert_eq!(
        pivot1.row_fields()[0].sort_type(),
        PivotFieldSort::Ascending
    );
    assert_eq!(pivot1.data_fields().len(), 1);
    assert_eq!(pivot1.data_fields()[0].custom_name(), Some("Total Sales"));

    // Verify second pivot table
    let pivot2 = &sheet.pivot_tables()[1];
    assert_eq!(pivot2.name(), "UnitsByProduct");
    assert_eq!(pivot2.target_row(), 12);
    assert_eq!(pivot2.target_col(), 5);
    assert_eq!(pivot2.row_fields()[0].name(), "Product");
    assert_eq!(
        pivot2.row_fields()[0].sort_type(),
        PivotFieldSort::Descending
    );
    assert_eq!(pivot2.data_fields()[0].custom_name(), Some("Total Units"));

    // Verify third pivot table
    let pivot3 = &sheet.pivot_tables()[2];
    assert_eq!(pivot3.name(), "CategoryBreakdown");
    assert_eq!(pivot3.target_row(), 12);
    assert_eq!(pivot3.target_col(), 10);
    assert_eq!(pivot3.row_fields()[0].name(), "Category");
    assert_eq!(
        pivot3.data_fields()[0].subtotal(),
        PivotSubtotalFunction::Average
    );

    // Verify fourth pivot table (complex)
    let pivot4 = &sheet.pivot_tables()[3];
    assert_eq!(pivot4.name(), "ComplexPivot");
    assert_eq!(pivot4.row_fields().len(), 2);
    assert_eq!(pivot4.row_fields()[0].name(), "Region");
    assert_eq!(pivot4.row_fields()[1].name(), "Category");
}

#[test]
fn test_pivot_table_clear() {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_sheet("ClearTest");

    // Add pivot tables
    for i in 1..=3 {
        let source = PivotSourceReference::from_range("Data!$A$1:$C$10");
        let pivot = PivotTable::new(format!("Pivot{}", i), source);
        sheet.add_pivot_table(pivot);
    }

    assert_eq!(sheet.pivot_tables().len(), 3);

    // Clear all pivot tables
    sheet.clear_pivot_tables();
    assert_eq!(sheet.pivot_tables().len(), 0);
}

#[test]
fn test_comprehensive_file_roundtrip_with_relationships() {
    let temp_dir = TempDir::new().expect("failed to create temp dir");
    let file_path = temp_dir.path().join("comprehensive_test.xlsx");

    // Create workbook with complex pivot table setup
    {
        let mut workbook = Workbook::new();
        create_sample_sales_data(&mut workbook, "SalesData");

        // Create multiple sheets with pivot tables
        let data_sheet = workbook.sheet_mut("SalesData").unwrap();

        // Pivot 1: Basic sales analysis
        let source1 = PivotSourceReference::from_range("SalesData!$A$1:$F$9");
        let mut pivot1 = PivotTable::new("SalesAnalysis", source1);
        pivot1
            .set_target(12, 0)
            .set_show_row_grand_totals(true)
            .set_show_column_grand_totals(true)
            .set_row_header_caption("Regions")
            .set_column_header_caption("Products")
            .set_preserve_formatting(true)
            .set_use_auto_formatting(false);

        let mut region_field = PivotField::new("Region");
        region_field
            .set_sort_type(PivotFieldSort::Ascending)
            .set_show_all_subtotals(true);
        pivot1.add_row_field(region_field);

        let mut product_field = PivotField::new("Product");
        product_field.set_sort_type(PivotFieldSort::Descending);
        pivot1.add_column_field(product_field);

        let mut sales_sum = PivotDataField::new("Sales");
        sales_sum
            .set_custom_name("Total Sales")
            .set_subtotal(PivotSubtotalFunction::Sum)
            .set_number_format("#,##0.00");
        pivot1.add_data_field(sales_sum);

        let mut sales_avg = PivotDataField::new("Sales");
        sales_avg
            .set_custom_name("Avg Sales")
            .set_subtotal(PivotSubtotalFunction::Average)
            .set_number_format("0.00");
        pivot1.add_data_field(sales_avg);

        data_sheet.add_pivot_table(pivot1);

        // Pivot 2: Units analysis with filters
        let source2 = PivotSourceReference::from_range("SalesData!$A$1:$F$9");
        let mut pivot2 = PivotTable::new("UnitsAnalysis", source2);
        pivot2
            .set_target(25, 0)
            .set_show_row_grand_totals(false)
            .set_show_column_grand_totals(true);

        let mut category_field = PivotField::new("Category");
        category_field
            .set_custom_label("Product Category")
            .set_show_empty_items(true);
        pivot2.add_row_field(category_field);

        let mut region_filter = PivotField::new("Region");
        region_filter.set_sort_type(PivotFieldSort::Manual);
        pivot2.add_page_field(region_filter);

        let mut units_count = PivotDataField::new("Units");
        units_count
            .set_custom_name("Total Units")
            .set_subtotal(PivotSubtotalFunction::Sum);
        pivot2.add_data_field(units_count);

        let mut units_max = PivotDataField::new("Units");
        units_max
            .set_custom_name("Max Units")
            .set_subtotal(PivotSubtotalFunction::Max);
        pivot2.add_data_field(units_max);

        let mut units_min = PivotDataField::new("Units");
        units_min
            .set_custom_name("Min Units")
            .set_subtotal(PivotSubtotalFunction::Min);
        pivot2.add_data_field(units_min);

        data_sheet.add_pivot_table(pivot2);

        // Add summary sheet
        let summary_sheet = workbook.add_sheet("Summary");
        summary_sheet
            .cell_mut("A1")
            .unwrap()
            .set_value("Analysis Summary");
        summary_sheet
            .cell_mut("A2")
            .unwrap()
            .set_value("See SalesData sheet for pivot tables");

        workbook.save(&file_path).expect("failed to save workbook");
    }

    // Verify file exists and has expected size
    assert!(file_path.exists(), "workbook file should exist");
    let metadata = std::fs::metadata(&file_path).expect("failed to get metadata");
    assert!(metadata.len() > 1024, "file should have content");

    // Reload and verify complete preservation
    let workbook = Workbook::open(&file_path).expect("failed to open workbook");

    // Verify sheets exist
    assert_eq!(workbook.worksheets().len(), 2);
    let data_sheet = workbook
        .sheet("SalesData")
        .expect("SalesData sheet should exist");
    let summary_sheet = workbook
        .sheet("Summary")
        .expect("Summary sheet should exist");

    // Verify data preserved
    assert_eq!(
        data_sheet.cell("A1").unwrap().value(),
        Some(&CellValue::String("Region".to_string()))
    );
    assert_eq!(
        data_sheet.cell("B1").unwrap().value(),
        Some(&CellValue::String("Product".to_string()))
    );
    assert_eq!(
        summary_sheet.cell("A1").unwrap().value(),
        Some(&CellValue::String("Analysis Summary".to_string()))
    );

    // Verify pivot tables count
    assert_eq!(data_sheet.pivot_tables().len(), 2);
    assert_eq!(summary_sheet.pivot_tables().len(), 0);

    // Verify Pivot 1 in detail
    let pivot1 = &data_sheet.pivot_tables()[0];
    assert_eq!(pivot1.name(), "SalesAnalysis");
    assert_eq!(pivot1.target_row(), 12);
    assert_eq!(pivot1.target_col(), 0);
    assert_eq!(pivot1.show_row_grand_totals(), true);
    assert_eq!(pivot1.show_column_grand_totals(), true);
    assert_eq!(pivot1.row_header_caption(), Some("Regions"));
    assert_eq!(pivot1.column_header_caption(), Some("Products"));
    assert_eq!(pivot1.preserve_formatting(), true);
    assert_eq!(pivot1.use_auto_formatting(), false);

    assert_eq!(pivot1.row_fields().len(), 1);
    assert_eq!(pivot1.row_fields()[0].name(), "Region");
    assert_eq!(
        pivot1.row_fields()[0].sort_type(),
        PivotFieldSort::Ascending
    );
    assert_eq!(pivot1.row_fields()[0].show_all_subtotals(), true);

    assert_eq!(pivot1.column_fields().len(), 1);
    assert_eq!(pivot1.column_fields()[0].name(), "Product");
    assert_eq!(
        pivot1.column_fields()[0].sort_type(),
        PivotFieldSort::Descending
    );

    assert_eq!(pivot1.data_fields().len(), 2);
    assert_eq!(pivot1.data_fields()[0].field_name(), "Sales");
    assert_eq!(pivot1.data_fields()[0].custom_name(), Some("Total Sales"));
    assert_eq!(
        pivot1.data_fields()[0].subtotal(),
        PivotSubtotalFunction::Sum
    );
    assert_eq!(pivot1.data_fields()[0].number_format(), Some("#,##0.00"));

    assert_eq!(pivot1.data_fields()[1].field_name(), "Sales");
    assert_eq!(pivot1.data_fields()[1].custom_name(), Some("Avg Sales"));
    assert_eq!(
        pivot1.data_fields()[1].subtotal(),
        PivotSubtotalFunction::Average
    );
    assert_eq!(pivot1.data_fields()[1].number_format(), Some("0.00"));

    // Verify Pivot 2 in detail
    let pivot2 = &data_sheet.pivot_tables()[1];
    assert_eq!(pivot2.name(), "UnitsAnalysis");
    assert_eq!(pivot2.target_row(), 25);
    assert_eq!(pivot2.target_col(), 0);
    assert_eq!(pivot2.show_row_grand_totals(), false);
    assert_eq!(pivot2.show_column_grand_totals(), true);

    assert_eq!(pivot2.row_fields().len(), 1);
    assert_eq!(pivot2.row_fields()[0].name(), "Category");
    assert_eq!(
        pivot2.row_fields()[0].custom_label(),
        Some("Product Category")
    );
    assert_eq!(pivot2.row_fields()[0].show_empty_items(), true);

    assert_eq!(pivot2.page_fields().len(), 1);
    assert_eq!(pivot2.page_fields()[0].name(), "Region");
    assert_eq!(pivot2.page_fields()[0].sort_type(), PivotFieldSort::Manual);

    assert_eq!(pivot2.data_fields().len(), 3);
    assert_eq!(pivot2.data_fields()[0].custom_name(), Some("Total Units"));
    assert_eq!(
        pivot2.data_fields()[0].subtotal(),
        PivotSubtotalFunction::Sum
    );
    assert_eq!(pivot2.data_fields()[1].custom_name(), Some("Max Units"));
    assert_eq!(
        pivot2.data_fields()[1].subtotal(),
        PivotSubtotalFunction::Max
    );
    assert_eq!(pivot2.data_fields()[2].custom_name(), Some("Min Units"));
    assert_eq!(
        pivot2.data_fields()[2].subtotal(),
        PivotSubtotalFunction::Min
    );

    // Verify source references preserved
    match pivot1.source_reference() {
        PivotSourceReference::WorksheetRange(range) => {
            assert_eq!(range, "SalesData!$A$1:$F$9");
        }
        _ => panic!("Expected WorksheetRange source"),
    }

    match pivot2.source_reference() {
        PivotSourceReference::WorksheetRange(range) => {
            assert_eq!(range, "SalesData!$A$1:$F$9");
        }
        _ => panic!("Expected WorksheetRange source"),
    }
}

#[test]
fn test_pivot_table_field_customization() {
    let temp_dir = TempDir::new().expect("failed to create temp dir");
    let file_path = temp_dir.path().join("field_custom_test.xlsx");

    // Create workbook with heavily customized fields
    {
        let mut workbook = Workbook::new();
        create_sample_sales_data(&mut workbook, "Data");
        let sheet = workbook.sheet_mut("Data").unwrap();

        let source = PivotSourceReference::from_range("Data!$A$1:$F$9");
        let mut pivot = PivotTable::new("CustomizedPivot", source);
        pivot.set_target(15, 0);

        // Row field with all customizations
        let mut row_field = PivotField::new("Region");
        row_field
            .set_custom_label("Geographic Region")
            .set_show_all_subtotals(true)
            .set_insert_blank_rows(true)
            .set_show_empty_items(true)
            .set_sort_type(PivotFieldSort::Ascending)
            .set_insert_page_breaks(false);
        pivot.add_row_field(row_field);

        // Data field with all customizations
        let mut data_field = PivotDataField::new("Sales");
        data_field
            .set_custom_name("Total Revenue (USD)")
            .set_subtotal(PivotSubtotalFunction::Sum)
            .set_number_format("$#,##0.00");
        pivot.add_data_field(data_field);

        sheet.add_pivot_table(pivot);

        workbook.save(&file_path).expect("failed to save workbook");
    }

    // Reload and verify all customizations
    let workbook = Workbook::open(&file_path).expect("failed to open workbook");
    let sheet = workbook.sheet("Data").expect("sheet should exist");

    let pivot = &sheet.pivot_tables()[0];
    let row_field = &pivot.row_fields()[0];

    assert_eq!(row_field.name(), "Region");
    assert_eq!(row_field.custom_label(), Some("Geographic Region"));
    assert_eq!(row_field.show_all_subtotals(), true);
    assert_eq!(row_field.insert_blank_rows(), true);
    assert_eq!(row_field.show_empty_items(), true);
    assert_eq!(row_field.sort_type(), PivotFieldSort::Ascending);
    assert_eq!(row_field.insert_page_breaks(), false);

    let data_field = &pivot.data_fields()[0];
    assert_eq!(data_field.field_name(), "Sales");
    assert_eq!(data_field.custom_name(), Some("Total Revenue (USD)"));
    assert_eq!(data_field.subtotal(), PivotSubtotalFunction::Sum);
    assert_eq!(data_field.number_format(), Some("$#,##0.00"));
}

#[test]
fn test_all_aggregation_functions_roundtrip() {
    let temp_dir = TempDir::new().expect("failed to create temp dir");
    let file_path = temp_dir.path().join("all_functions_test.xlsx");

    // Create workbook with pivot table using all aggregation functions
    {
        let mut workbook = Workbook::new();
        create_sample_sales_data(&mut workbook, "Data");
        let sheet = workbook.sheet_mut("Data").unwrap();

        let source = PivotSourceReference::from_range("Data!$A$1:$F$9");
        let mut pivot = PivotTable::new("AllSubtotals", source);
        pivot.set_target(0, 10);

        // Add row field
        let region_field = PivotField::new("Region");
        pivot.add_row_field(region_field);

        // Add a data field for each subtotal function
        let functions = vec![
            (PivotSubtotalFunction::Sum, "Sum of Sales"),
            (PivotSubtotalFunction::Average, "Average Sales"),
            (PivotSubtotalFunction::Count, "Count Sales"),
            (PivotSubtotalFunction::CountNums, "CountNums Sales"),
            (PivotSubtotalFunction::Max, "Max Sales"),
            (PivotSubtotalFunction::Min, "Min Sales"),
            (PivotSubtotalFunction::Product, "Product Sales"),
            (PivotSubtotalFunction::StdDev, "StdDev Sales"),
            (PivotSubtotalFunction::StdDevP, "StdDevP Sales"),
            (PivotSubtotalFunction::Var, "Var Sales"),
            (PivotSubtotalFunction::VarP, "VarP Sales"),
        ];

        for (func, name) in functions {
            let mut data_field = PivotDataField::new("Sales");
            data_field.set_custom_name(name).set_subtotal(func);
            pivot.add_data_field(data_field);
        }

        sheet.add_pivot_table(pivot);

        workbook.save(&file_path).expect("failed to save workbook");
    }

    // Reload and verify all functions preserved
    let workbook = Workbook::open(&file_path).expect("failed to open workbook");
    let sheet = workbook.sheet("Data").expect("sheet should exist");

    assert_eq!(sheet.pivot_tables().len(), 1);
    let pivot = &sheet.pivot_tables()[0];
    assert_eq!(pivot.data_fields().len(), 11);

    // Verify each function
    let expected_functions = vec![
        (PivotSubtotalFunction::Sum, "Sum of Sales"),
        (PivotSubtotalFunction::Average, "Average Sales"),
        (PivotSubtotalFunction::Count, "Count Sales"),
        (PivotSubtotalFunction::CountNums, "CountNums Sales"),
        (PivotSubtotalFunction::Max, "Max Sales"),
        (PivotSubtotalFunction::Min, "Min Sales"),
        (PivotSubtotalFunction::Product, "Product Sales"),
        (PivotSubtotalFunction::StdDev, "StdDev Sales"),
        (PivotSubtotalFunction::StdDevP, "StdDevP Sales"),
        (PivotSubtotalFunction::Var, "Var Sales"),
        (PivotSubtotalFunction::VarP, "VarP Sales"),
    ];

    for (idx, (expected_func, expected_name)) in expected_functions.iter().enumerate() {
        let data_field = &pivot.data_fields()[idx];
        assert_eq!(data_field.field_name(), "Sales");
        assert_eq!(data_field.custom_name(), Some(*expected_name));
        assert_eq!(data_field.subtotal(), *expected_func);
    }
}

#[test]
fn test_complex_pivot_configuration_roundtrip() {
    let temp_dir = TempDir::new().expect("failed to create temp dir");
    let file_path = temp_dir.path().join("complex_pivot_test.xlsx");

    // Create workbook with complex pivot table configuration
    {
        let mut workbook = Workbook::new();
        create_sample_sales_data(&mut workbook, "Data");
        let sheet = workbook.sheet_mut("Data").unwrap();

        let source = PivotSourceReference::from_range("Data!$A$1:$F$9");
        let mut pivot = PivotTable::new("ComplexPivot", source);
        pivot
            .set_target(15, 0)
            .set_show_row_grand_totals(true)
            .set_show_column_grand_totals(false)
            .set_row_header_caption("Custom Row Header")
            .set_column_header_caption("Custom Column Header");

        // Multiple row fields with different configurations
        let mut region_field = PivotField::new("Region");
        region_field
            .set_sort_type(PivotFieldSort::Ascending)
            .set_show_all_subtotals(true)
            .set_insert_blank_rows(false);
        pivot.add_row_field(region_field);

        let mut product_field = PivotField::new("Product");
        product_field
            .set_custom_label("Product Name")
            .set_sort_type(PivotFieldSort::Descending)
            .set_show_empty_items(true);
        pivot.add_row_field(product_field);

        let mut category_field = PivotField::new("Category");
        category_field
            .set_sort_type(PivotFieldSort::Manual)
            .set_insert_page_breaks(false);
        pivot.add_row_field(category_field);

        // Multiple column fields
        let mut col_field1 = PivotField::new("Units");
        col_field1.set_sort_type(PivotFieldSort::Ascending);
        pivot.add_column_field(col_field1);

        let mut col_field2 = PivotField::new("Price");
        col_field2.set_custom_label("Unit Price");
        pivot.add_column_field(col_field2);

        // Page filters (filter fields)
        let mut page_filter = PivotField::new("Sales");
        page_filter.set_sort_type(PivotFieldSort::Descending);
        pivot.add_page_field(page_filter);

        // Multiple data fields with different aggregations
        let mut sales_sum = PivotDataField::new("Sales");
        sales_sum
            .set_custom_name("Total Sales")
            .set_subtotal(PivotSubtotalFunction::Sum)
            .set_number_format("#,##0.00");
        pivot.add_data_field(sales_sum);

        let mut sales_avg = PivotDataField::new("Sales");
        sales_avg
            .set_custom_name("Average Sales")
            .set_subtotal(PivotSubtotalFunction::Average)
            .set_number_format("0.00");
        pivot.add_data_field(sales_avg);

        let mut units_count = PivotDataField::new("Units");
        units_count
            .set_custom_name("Unit Count")
            .set_subtotal(PivotSubtotalFunction::CountNums);
        pivot.add_data_field(units_count);

        sheet.add_pivot_table(pivot);

        workbook.save(&file_path).expect("failed to save workbook");
    }

    // Reload and verify complex configuration
    let workbook = Workbook::open(&file_path).expect("failed to open workbook");
    let sheet = workbook.sheet("Data").expect("sheet should exist");

    assert_eq!(sheet.pivot_tables().len(), 1);
    let pivot = &sheet.pivot_tables()[0];

    // Verify pivot table properties
    assert_eq!(pivot.name(), "ComplexPivot");
    assert_eq!(pivot.target_row(), 15);
    assert_eq!(pivot.target_col(), 0);
    assert_eq!(pivot.show_row_grand_totals(), true);
    assert_eq!(pivot.show_column_grand_totals(), false);
    assert_eq!(pivot.row_header_caption(), Some("Custom Row Header"));
    assert_eq!(pivot.column_header_caption(), Some("Custom Column Header"));

    // Verify row fields (3 fields)
    assert_eq!(pivot.row_fields().len(), 3);

    let row1 = &pivot.row_fields()[0];
    assert_eq!(row1.name(), "Region");
    assert_eq!(row1.sort_type(), PivotFieldSort::Ascending);
    assert_eq!(row1.show_all_subtotals(), true);
    assert_eq!(row1.insert_blank_rows(), false);

    let row2 = &pivot.row_fields()[1];
    assert_eq!(row2.name(), "Product");
    assert_eq!(row2.custom_label(), Some("Product Name"));
    assert_eq!(row2.sort_type(), PivotFieldSort::Descending);
    assert_eq!(row2.show_empty_items(), true);

    let row3 = &pivot.row_fields()[2];
    assert_eq!(row3.name(), "Category");
    assert_eq!(row3.sort_type(), PivotFieldSort::Manual);
    assert_eq!(row3.insert_page_breaks(), false);

    // Verify column fields (2 fields)
    assert_eq!(pivot.column_fields().len(), 2);

    let col1 = &pivot.column_fields()[0];
    assert_eq!(col1.name(), "Units");
    assert_eq!(col1.sort_type(), PivotFieldSort::Ascending);

    let col2 = &pivot.column_fields()[1];
    assert_eq!(col2.name(), "Price");
    assert_eq!(col2.custom_label(), Some("Unit Price"));

    // Verify page fields (1 field)
    assert_eq!(pivot.page_fields().len(), 1);
    assert_eq!(pivot.page_fields()[0].name(), "Sales");
    assert_eq!(
        pivot.page_fields()[0].sort_type(),
        PivotFieldSort::Descending
    );

    // Verify data fields (3 fields)
    assert_eq!(pivot.data_fields().len(), 3);

    let data1 = &pivot.data_fields()[0];
    assert_eq!(data1.field_name(), "Sales");
    assert_eq!(data1.custom_name(), Some("Total Sales"));
    assert_eq!(data1.subtotal(), PivotSubtotalFunction::Sum);
    assert_eq!(data1.number_format(), Some("#,##0.00"));

    let data2 = &pivot.data_fields()[1];
    assert_eq!(data2.field_name(), "Sales");
    assert_eq!(data2.custom_name(), Some("Average Sales"));
    assert_eq!(data2.subtotal(), PivotSubtotalFunction::Average);
    assert_eq!(data2.number_format(), Some("0.00"));

    let data3 = &pivot.data_fields()[2];
    assert_eq!(data3.field_name(), "Units");
    assert_eq!(data3.custom_name(), Some("Unit Count"));
    assert_eq!(data3.subtotal(), PivotSubtotalFunction::CountNums);
}

#[test]
fn test_pivot_source_reference_types() {
    let temp_dir = TempDir::new().expect("failed to create temp dir");
    let file_path = temp_dir.path().join("source_types_test.xlsx");

    // Create workbook with different source reference types
    {
        let mut workbook = Workbook::new();
        create_sample_sales_data(&mut workbook, "Data");
        let sheet = workbook.sheet_mut("Data").unwrap();

        // Test WorksheetRange source
        let source_range = PivotSourceReference::from_range("Data!$A$1:$F$9");
        let mut pivot_range = PivotTable::new("RangePivot", source_range);
        pivot_range.set_target(15, 0);
        let region_field = PivotField::new("Region");
        pivot_range.add_row_field(region_field);
        let sales_data = PivotDataField::new("Sales");
        pivot_range.add_data_field(sales_data);
        sheet.add_pivot_table(pivot_range);

        // Test NamedTable source
        let source_table = PivotSourceReference::from_table("SalesTable");
        let mut pivot_table = PivotTable::new("TablePivot", source_table);
        pivot_table.set_target(15, 5);
        let product_field = PivotField::new("Product");
        pivot_table.add_row_field(product_field);
        let units_data = PivotDataField::new("Units");
        pivot_table.add_data_field(units_data);
        sheet.add_pivot_table(pivot_table);

        workbook.save(&file_path).expect("failed to save workbook");
    }

    // Reload and verify source references
    let workbook = Workbook::open(&file_path).expect("failed to open workbook");
    let sheet = workbook.sheet("Data").expect("sheet should exist");

    assert_eq!(sheet.pivot_tables().len(), 2);

    // Verify WorksheetRange source
    let pivot1 = &sheet.pivot_tables()[0];
    assert_eq!(pivot1.name(), "RangePivot");
    match pivot1.source_reference() {
        PivotSourceReference::WorksheetRange(range) => {
            assert_eq!(range, "Data!$A$1:$F$9");
        }
        _ => panic!("Expected WorksheetRange source"),
    }

    // Verify NamedTable source
    let pivot2 = &sheet.pivot_tables()[1];
    assert_eq!(pivot2.name(), "TablePivot");
    match pivot2.source_reference() {
        PivotSourceReference::NamedTable(table_name) => {
            assert_eq!(table_name, "SalesTable");
        }
        _ => panic!("Expected NamedTable source"),
    }
}

#[test]
fn test_worksheet_pivot_table_integration() {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_sheet("Integration");

    // Test accessing empty pivot tables
    assert_eq!(sheet.pivot_tables().len(), 0);

    // Add pivot tables
    for i in 1..=5 {
        let source = PivotSourceReference::from_range(format!("Sheet!$A$1:$D${}", i * 10));
        let mut pivot = PivotTable::new(format!("Pivot{}", i), source);
        pivot.set_target(i * 5, 0);
        sheet.add_pivot_table(pivot);
    }

    // Verify count
    assert_eq!(sheet.pivot_tables().len(), 5);

    // Access by index
    for i in 0..5 {
        let pivot = &sheet.pivot_tables()[i];
        assert_eq!(pivot.name(), format!("Pivot{}", i + 1));
        assert_eq!(pivot.target_row(), ((i + 1) * 5) as u32);
    }

    // Clear all pivot tables
    sheet.clear_pivot_tables();
    assert_eq!(sheet.pivot_tables().len(), 0);

    // Add one more after clearing
    let source = PivotSourceReference::from_range("Sheet!$A$1:$D$10");
    let pivot = PivotTable::new("NewPivot", source);
    sheet.add_pivot_table(pivot);
    assert_eq!(sheet.pivot_tables().len(), 1);
    assert_eq!(sheet.pivot_tables()[0].name(), "NewPivot");
}

#[test]
fn test_pivot_field_sort_types() {
    let temp_dir = TempDir::new().expect("failed to create temp dir");
    let file_path = temp_dir.path().join("sort_test.xlsx");

    // Create workbook with all sort types
    {
        let mut workbook = Workbook::new();
        create_sample_sales_data(&mut workbook, "Data");
        let sheet = workbook.sheet_mut("Data").unwrap();

        let source = PivotSourceReference::from_range("Data!$A$1:$F$9");
        let mut pivot = PivotTable::new("SortedPivot", source);
        pivot.set_target(15, 0);

        // Manual sort
        let mut manual_field = PivotField::new("Region");
        manual_field.set_sort_type(PivotFieldSort::Manual);
        pivot.add_row_field(manual_field);

        // Ascending sort
        let mut asc_field = PivotField::new("Product");
        asc_field.set_sort_type(PivotFieldSort::Ascending);
        pivot.add_column_field(asc_field);

        // Descending sort
        let mut desc_field = PivotField::new("Category");
        desc_field.set_sort_type(PivotFieldSort::Descending);
        pivot.add_page_field(desc_field);

        sheet.add_pivot_table(pivot);

        workbook.save(&file_path).expect("failed to save workbook");
    }

    // Reload and verify sort types
    let workbook = Workbook::open(&file_path).expect("failed to open workbook");
    let sheet = workbook.sheet("Data").expect("sheet should exist");

    let pivot = &sheet.pivot_tables()[0];
    assert_eq!(pivot.row_fields()[0].sort_type(), PivotFieldSort::Manual);
    assert_eq!(
        pivot.column_fields()[0].sort_type(),
        PivotFieldSort::Ascending
    );
    assert_eq!(
        pivot.page_fields()[0].sort_type(),
        PivotFieldSort::Descending
    );
}
