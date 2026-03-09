use std::path::Path;

use anyhow::{bail, Result};

use crate::cmd::set::resolve_output;
use crate::format::FileFormat;

pub fn run(
    path: &Path,
    find: &str,
    replace_with: &str,
    output_path: Option<&Path>,
    in_place: bool,
) -> Result<()> {
    let dest = resolve_output(path, output_path, in_place)?;
    let format = FileFormat::detect(path)?;
    match format {
        FileFormat::Xlsx => replace_xlsx(path, find, replace_with, &dest),
        FileFormat::Docx => replace_docx(path, find, replace_with, &dest),
        FileFormat::Pptx => bail!("pptx replace is not yet implemented"),
    }
}

fn replace_xlsx(path: &Path, find: &str, replace_with: &str, dest: &Path) -> Result<()> {
    let mut wb = offidized_xlsx::Workbook::open(path)?;

    for ws in wb.worksheets_mut() {
        let refs_to_update: Vec<String> = ws
            .cells()
            .filter_map(|(cell_ref, cell)| {
                if let Some(offidized_xlsx::CellValue::String(s)) = cell.value() {
                    if s.contains(find) {
                        return Some(cell_ref.to_string());
                    }
                }
                None
            })
            .collect();

        for cell_ref in &refs_to_update {
            let cell = ws.cell_mut(cell_ref)?;
            if let Some(offidized_xlsx::CellValue::String(s)) = cell.value().cloned() {
                let new_val = s.replace(find, replace_with);
                cell.set_value(offidized_xlsx::CellValue::String(new_val));
            }
        }
    }

    wb.save(dest)?;
    Ok(())
}

fn replace_docx(path: &Path, find: &str, replace_with: &str, dest: &Path) -> Result<()> {
    let mut doc = offidized_docx::Document::open(path)?;

    for paragraph in doc.paragraphs_mut() {
        for run in paragraph.runs_mut() {
            let text = run.text().to_string();
            if text.contains(find) {
                run.set_text(text.replace(find, replace_with));
            }
        }
    }

    doc.save(dest)?;
    Ok(())
}
