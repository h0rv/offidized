//! Example demonstrating formula evaluation in offidized-xlsx.
//!
//! This example shows how to:
//! 1. Create a workbook with formulas
//! 2. Evaluate formulas to get computed values
//! 3. Use the formula engine for various Excel functions

use offidized_xlsx::{CellValue, Workbook};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create a new workbook
    let mut wb = Workbook::new();

    // Add data and formulas
    {
        let ws = wb.add_sheet("Sheet1");

        // Simple arithmetic
        ws.cell_mut("A1")?.set_value(10);
        ws.cell_mut("A2")?.set_value(20);
        ws.cell_mut("A3")?.set_formula("A1+A2");

        // SUM function
        ws.cell_mut("B1")?.set_value(5);
        ws.cell_mut("B2")?.set_value(15);
        ws.cell_mut("B3")?.set_value(25);
        ws.cell_mut("B4")?.set_formula("SUM(B1:B3)");

        // More complex formula
        ws.cell_mut("C1")?.set_value(100);
        ws.cell_mut("C2")?.set_formula("SQRT(C1)");

        // Text concatenation
        ws.cell_mut("D1")?.set_value("Hello");
        ws.cell_mut("D2")?.set_value("World");
        ws.cell_mut("D3")?.set_formula("CONCATENATE(D1,\" \",D2)");

        // Cross-cell references
        ws.cell_mut("E1")?.set_formula("A1*B1");
        ws.cell_mut("E2")?.set_formula("E1+10");
    }

    // Evaluate formulas
    let ws = wb.sheet("Sheet1").ok_or("Sheet not found")?;

    println!("Formula Evaluation Results:");
    println!("===========================");

    // Simple arithmetic
    let result = ws.evaluate_formula("A3", &wb)?;
    println!("A3 (=A1+A2): {:?}", result);
    assert_eq!(result, CellValue::Number(30.0));

    // SUM function
    let result = ws.evaluate_formula("B4", &wb)?;
    println!("B4 (=SUM(B1:B3)): {:?}", result);
    assert_eq!(result, CellValue::Number(45.0));

    // SQRT function
    let result = ws.evaluate_formula("C2", &wb)?;
    println!("C2 (=SQRT(C1)): {:?}", result);
    assert_eq!(result, CellValue::Number(10.0));

    // Text concatenation
    let result = ws.evaluate_formula("D3", &wb)?;
    println!("D3 (=CONCATENATE(D1,\" \",D2)): {:?}", result);
    assert_eq!(result, CellValue::String("Hello World".to_string()));

    // Cross-cell references
    let result = ws.evaluate_formula("E2", &wb)?;
    println!("E2 (=E1+10 where E1=A1*B1): {:?}", result);
    assert_eq!(result, CellValue::Number(60.0)); // (10*5)+10 = 60

    println!("\n✓ All formulas evaluated successfully!");

    Ok(())
}
