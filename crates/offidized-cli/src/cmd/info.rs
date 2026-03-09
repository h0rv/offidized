use std::path::Path;

use anyhow::{bail, Result};
use offidized_opc::Package;

use crate::format::FileFormat;
use crate::output::{DefinedNameInfo, DocxInfo, XlsxInfo};

pub fn run(path: &Path) -> Result<()> {
    let format = FileFormat::detect(path)?;
    match format {
        FileFormat::Xlsx => info_xlsx(path),
        FileFormat::Docx => info_docx(path),
        FileFormat::Pptx => bail!("pptx info is not yet implemented"),
    }
}

fn info_xlsx(path: &Path) -> Result<()> {
    let wb = offidized_xlsx::Workbook::open(path)?;
    let sheets: Vec<String> = wb.sheet_names().into_iter().map(String::from).collect();
    let defined_names: Vec<DefinedNameInfo> = wb
        .defined_names()
        .iter()
        .map(|dn| DefinedNameInfo {
            name: dn.name().to_string(),
            reference: dn.reference().to_string(),
        })
        .collect();
    // Workbook doesn't expose the underlying Package, so re-open as OPC for part count.
    // TODO: expose source_package() on Workbook to avoid the double-open.
    let part_count = Package::open(path).map(|p| p.part_count()).unwrap_or(0);
    let info = XlsxInfo {
        format: "xlsx",
        sheets,
        defined_names,
        part_count,
    };
    println!("{}", serde_json::to_string(&info)?);
    Ok(())
}

fn info_docx(path: &Path) -> Result<()> {
    let doc = offidized_docx::Document::open(path)?;
    let paragraph_count = doc.paragraphs().len();
    let table_count = doc.tables().len();
    let part_count = doc.package().part_count();
    let info = DocxInfo {
        format: "docx",
        paragraph_count,
        table_count,
        part_count,
    };
    println!("{}", serde_json::to_string(&info)?);
    Ok(())
}
