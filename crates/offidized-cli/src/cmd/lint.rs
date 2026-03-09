use std::path::Path;

use offidized_xlsx::Workbook;

pub fn run(file: &Path) -> anyhow::Result<()> {
    let workbook = Workbook::open(file)?;
    let report = workbook
        .lint()
        .check_broken_refs()
        .check_formula_consistency()
        .check_pivot_sources()
        .check_named_ranges()
        .run();

    println!("{{");
    println!("  \"errors\": {},", report.error_count());
    println!("  \"warnings\": {},", report.warning_count());
    println!("  \"findings\": [");
    for (index, finding) in report.findings().iter().enumerate() {
        let comma = if index + 1 == report.findings().len() {
            ""
        } else {
            ","
        };
        println!(
            "    {{\"severity\":\"{:?}\",\"code\":\"{}\",\"message\":\"{}\",\"sheet\":{},\"cell\":{},\"object\":{}}}{}",
            finding.severity,
            finding.code,
            escape_json(finding.message.as_str()),
            finding
                .location
                .sheet
                .as_ref()
                .map(|v| format!("\"{}\"", escape_json(v)))
                .unwrap_or_else(|| "null".to_string()),
            finding
                .location
                .cell
                .as_ref()
                .map(|v| format!("\"{}\"", escape_json(v)))
                .unwrap_or_else(|| "null".to_string()),
            finding
                .location
                .object
                .as_ref()
                .map(|v| format!("\"{}\"", escape_json(v)))
                .unwrap_or_else(|| "null".to_string()),
            comma
        );
    }
    println!("  ]");
    println!("}}");

    Ok(())
}

fn escape_json(input: &str) -> String {
    input
        .replace('\\', "\\\\")
        .replace('"', "\\\"")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
        .replace('\t', "\\t")
}
