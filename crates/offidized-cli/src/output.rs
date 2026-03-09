use serde::Serialize;

// --- info ---

#[derive(Serialize)]
pub struct XlsxInfo {
    pub format: &'static str,
    pub sheets: Vec<String>,
    pub defined_names: Vec<DefinedNameInfo>,
    pub part_count: usize,
}

#[derive(Serialize)]
pub struct DefinedNameInfo {
    pub name: String,
    pub reference: String,
}

#[derive(Serialize)]
pub struct DocxInfo {
    pub format: &'static str,
    pub paragraph_count: usize,
    pub table_count: usize,
    pub part_count: usize,
}

// --- read ---

#[derive(Serialize)]
pub struct CellOutput {
    #[serde(rename = "ref")]
    pub cell_ref: String,
    pub value: serde_json::Value,
    #[serde(rename = "type")]
    pub value_type: String,
}

#[derive(Serialize)]
pub struct SheetCells {
    pub name: String,
    pub cells: Vec<CellOutput>,
}

#[derive(Serialize)]
pub struct ParagraphOutput {
    pub index: usize,
    pub text: String,
    pub style: Option<String>,
}

// --- part ---

#[derive(Serialize)]
pub struct PartInfo {
    pub uri: String,
    pub content_type: Option<String>,
    pub size_bytes: usize,
}
