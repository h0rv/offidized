use std::path::Path;

use anyhow::Result;
use offidized_ir::{UnifiedDeriveOptions, UnifiedDocument};
use serde::Serialize;

#[derive(Debug, Serialize)]
struct CapabilitiesOut {
    text_nodes: bool,
    table_cells: bool,
    chart_meta: bool,
    style_nodes: bool,
}

pub fn run(file: &Path) -> Result<()> {
    let doc = UnifiedDocument::derive(file, UnifiedDeriveOptions::default())?;
    let caps = doc.capabilities();
    let out = CapabilitiesOut {
        text_nodes: caps.text_nodes,
        table_cells: caps.table_cells,
        chart_meta: caps.chart_meta,
        style_nodes: caps.style_nodes,
    };
    println!("{}", serde_json::to_string_pretty(&out)?);
    Ok(())
}
