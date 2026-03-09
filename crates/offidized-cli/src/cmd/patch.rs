use std::io::Read as _;
use std::path::Path;

use anyhow::{bail, Result};
use serde::Deserialize;

use crate::cmd::set::resolve_output;
use crate::format::FileFormat;

#[derive(Deserialize)]
struct XlsxPatch {
    #[serde(rename = "ref")]
    cell_ref: String,
    value: Option<serde_json::Value>,
}

#[derive(Deserialize)]
struct DocxPatch {
    paragraph: usize,
    text: String,
}

pub fn run(path: &Path, output_path: Option<&Path>, in_place: bool) -> Result<()> {
    let dest = resolve_output(path, output_path, in_place)?;
    let format = FileFormat::detect(path)?;

    let mut stdin_buf = String::new();
    std::io::stdin().read_to_string(&mut stdin_buf)?;

    match format {
        FileFormat::Xlsx => patch_xlsx(path, &stdin_buf, &dest),
        FileFormat::Docx => patch_docx(path, &stdin_buf, &dest),
        FileFormat::Pptx => bail!("pptx patch is not yet implemented"),
    }
}

fn patch_xlsx(path: &Path, stdin: &str, dest: &Path) -> Result<()> {
    let patches: Vec<XlsxPatch> = serde_json::from_str(stdin)?;
    let mut wb = offidized_xlsx::Workbook::open(path)?;

    for patch in &patches {
        let (sheet_name, cell_ref) = patch.cell_ref.split_once('!').ok_or_else(|| {
            anyhow::anyhow!("patch ref must include sheet name: {:?}", patch.cell_ref)
        })?;

        let ws = wb
            .worksheets_mut()
            .iter_mut()
            .find(|ws| ws.name() == sheet_name)
            .ok_or_else(|| anyhow::anyhow!("sheet not found: {sheet_name:?}"))?;

        let cell = ws.cell_mut(cell_ref)?;
        match &patch.value {
            None | Some(serde_json::Value::Null) => {
                cell.clear_value();
            }
            Some(serde_json::Value::Number(n)) => {
                let f = n.as_f64().ok_or_else(|| {
                    anyhow::anyhow!("JSON number cannot be represented as f64: {n}")
                })?;
                cell.set_value(offidized_xlsx::CellValue::Number(f));
            }
            Some(serde_json::Value::Bool(b)) => {
                cell.set_value(offidized_xlsx::CellValue::Bool(*b));
            }
            Some(serde_json::Value::String(s)) => {
                cell.set_value(offidized_xlsx::CellValue::String(s.clone()));
            }
            Some(other) => {
                cell.set_value(offidized_xlsx::CellValue::String(other.to_string()));
            }
        }
    }

    wb.save(dest)?;
    Ok(())
}

fn patch_docx(path: &Path, stdin: &str, dest: &Path) -> Result<()> {
    let patches: Vec<DocxPatch> = serde_json::from_str(stdin)?;
    let mut doc = offidized_docx::Document::open(path)?;

    let paragraphs = doc.paragraphs_mut();
    for patch in &patches {
        if patch.paragraph >= paragraphs.len() {
            bail!(
                "paragraph index {} out of range (document has {} paragraphs)",
                patch.paragraph,
                paragraphs.len()
            );
        }
        paragraphs[patch.paragraph].set_text(&patch.text);
    }

    doc.save(dest)?;
    Ok(())
}
