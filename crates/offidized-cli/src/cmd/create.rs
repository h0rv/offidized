use std::path::Path;

use anyhow::{bail, Result};

use crate::format::FileFormat;

pub fn run(path: &Path, sheets: &[String]) -> Result<()> {
    let format = FileFormat::detect(path)?;
    match format {
        FileFormat::Xlsx => create_xlsx(path, sheets),
        FileFormat::Docx => create_docx(path),
        FileFormat::Pptx => bail!("pptx create is not yet implemented"),
    }
}

fn create_xlsx(path: &Path, sheets: &[String]) -> Result<()> {
    let mut wb = offidized_xlsx::Workbook::new();
    if sheets.is_empty() {
        wb.add_sheet("Sheet1");
    } else {
        for name in sheets {
            wb.add_sheet(name.as_str());
        }
    }
    wb.save(path)?;
    Ok(())
}

fn create_docx(path: &Path) -> Result<()> {
    let doc = offidized_docx::Document::new();
    doc.save(path)?;
    Ok(())
}
