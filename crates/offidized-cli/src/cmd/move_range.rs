use std::path::Path;

use anyhow::Result;
use offidized_xlsx::{range::CellRange, Workbook};

use crate::cmd::set::resolve_output;
use crate::range_parse::parse_range;

pub fn run(
    path: &Path,
    source_range: &str,
    dest_cell: &str,
    output_path: Option<&Path>,
    in_place: bool,
) -> Result<()> {
    let dest = resolve_output(path, output_path, in_place)?;

    let mut wb = Workbook::open(path)?;

    let source_parsed = parse_range(source_range)?;
    let dest_parsed = parse_range(dest_cell)?;

    let ws = wb
        .worksheets_mut()
        .iter_mut()
        .find(|ws| ws.name() == source_parsed.sheet)
        .ok_or_else(|| anyhow::anyhow!("source sheet not found: {:?}", source_parsed.sheet))?;

    let range = CellRange::parse(&source_parsed.range)?;
    range.move_to(ws, &dest_parsed.range)?;

    wb.save(&dest)?;
    Ok(())
}
