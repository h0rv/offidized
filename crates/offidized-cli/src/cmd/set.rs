use std::path::{Path, PathBuf};

use anyhow::{bail, Result};
use offidized_xlsx::CellValue;

use crate::format::FileFormat;
use crate::range_parse::parse_range;

pub fn run(
    path: &Path,
    target: &str,
    value: &str,
    output_path: Option<&Path>,
    in_place: bool,
) -> Result<()> {
    let dest = resolve_output(path, output_path, in_place)?;
    let format = FileFormat::detect(path)?;
    match format {
        FileFormat::Xlsx => set_xlsx(path, target, value, &dest),
        FileFormat::Docx => set_docx(path, target, value, &dest),
        FileFormat::Pptx => bail!("pptx set is not yet implemented"),
    }
}

fn set_xlsx(path: &Path, target: &str, value: &str, dest: &Path) -> Result<()> {
    let mut wb = offidized_xlsx::Workbook::open(path)?;
    let parsed = parse_range(target)?;
    let ws = wb
        .worksheets_mut()
        .iter_mut()
        .find(|ws| ws.name() == parsed.sheet)
        .ok_or_else(|| anyhow::anyhow!("sheet not found: {:?}", parsed.sheet))?;

    let cell = ws.cell_mut(&parsed.range)?;
    let cell_value = auto_type_value(value);
    cell.set_value(cell_value);
    wb.save(dest)?;
    Ok(())
}

fn set_docx(path: &Path, target: &str, value: &str, dest: &Path) -> Result<()> {
    let mut doc = offidized_docx::Document::open(path)?;
    let index: usize = target.parse().map_err(|_| {
        anyhow::anyhow!("docx set target must be a paragraph index (0-based), got: {target:?}")
    })?;

    let paragraphs = doc.paragraphs_mut();
    if index >= paragraphs.len() {
        bail!(
            "paragraph index {index} out of range (document has {} paragraphs)",
            paragraphs.len()
        );
    }
    paragraphs[index].set_text(value);
    doc.save(dest)?;
    Ok(())
}

/// Parse a string value into the most specific CellValue type.
pub fn auto_type_value(value: &str) -> CellValue {
    if let Ok(n) = value.parse::<f64>() {
        return CellValue::Number(n);
    }
    match value.to_ascii_lowercase().as_str() {
        "true" => CellValue::Bool(true),
        "false" => CellValue::Bool(false),
        _ => CellValue::String(value.to_string()),
    }
}

pub fn resolve_output(input: &Path, output: Option<&Path>, in_place: bool) -> Result<PathBuf> {
    match (output, in_place) {
        (Some(_), true) => bail!("cannot specify both -o and -i"),
        (Some(o), false) => Ok(o.to_path_buf()),
        (None, true) => Ok(input.to_path_buf()),
        (None, false) => bail!("must specify -o <path> or -i for in-place edit"),
    }
}
