use std::{
    collections::{HashMap, HashSet},
    fs,
    path::{Path, PathBuf},
};

use anyhow::{bail, Result};
use offidized_ir::{
    ApplyOptions, CellStylePatch, PptxTextStylePatch, UnifiedDeriveOptions, UnifiedDocument,
    UnifiedEdit, UnifiedEditPayload,
};
use serde::Deserialize;

#[derive(Deserialize)]
struct JsonEditSpec {
    #[serde(default)]
    file: Option<String>,
    id: String,
    #[serde(default)]
    text: Option<String>,
    #[serde(default)]
    group: Option<String>,
    #[serde(default)]
    payload: Option<JsonEditPayload>,
}

#[derive(Debug, Clone)]
struct ScopedEdit {
    file: PathBuf,
    edit: UnifiedEdit,
}

#[derive(Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
enum JsonEditPayload {
    XlsxCellStyle {
        #[serde(default)]
        bold: Option<bool>,
        #[serde(default)]
        italic: Option<bool>,
        #[serde(default)]
        number_format: Option<String>,
    },
    PptxTextStyle {
        #[serde(default)]
        bold: Option<bool>,
        #[serde(default)]
        italic: Option<bool>,
        #[serde(default)]
        font_size: Option<u32>,
        #[serde(default)]
        font_color: Option<String>,
        #[serde(default)]
        font_name: Option<String>,
    },
}

pub fn run(
    file: &Path,
    edit_specs: &[String],
    edits_json_path: Option<&Path>,
    output_path: Option<&Path>,
    in_place: bool,
    force: bool,
    strict: bool,
    lint: bool,
) -> Result<()> {
    let dest = crate::cmd::set::resolve_output(file, output_path, in_place)?;
    if edit_specs.is_empty() && edits_json_path.is_none() {
        bail!("at least one --edit '<id>=<text>' or --edits-json <path> is required");
    }

    let base_file = file.to_path_buf();
    let mut scoped_edits: Vec<ScopedEdit> = edit_specs
        .iter()
        .map(|spec| {
            Ok(ScopedEdit {
                file: base_file.clone(),
                edit: parse_edit_spec(spec)?,
            })
        })
        .collect::<Result<Vec<_>>>()?;
    if let Some(path) = edits_json_path {
        let json_edits = parse_edits_json(path, base_file.as_path())?;
        scoped_edits.extend(json_edits);
    }
    if scoped_edits.is_empty() {
        bail!("no edits were provided");
    }

    let unique_files: HashSet<PathBuf> = scoped_edits
        .iter()
        .map(|entry| entry.file.clone())
        .collect();
    if unique_files.len() > 1 && (!in_place || output_path.is_some()) {
        bail!("cross-file edit batches require --in-place and no --output");
    }

    let temp_dir = tempfile::tempdir()?;
    let mut shadow_paths: HashMap<PathBuf, PathBuf> = HashMap::new();
    for (index, file_path) in unique_files.iter().enumerate() {
        let shadow = temp_dir.path().join(format!("edit-{index}.tmp"));
        fs::copy(file_path, &shadow).map_err(|err| {
            anyhow::anyhow!(
                "failed to stage source file {} for transactional edits: {err}",
                file_path.display()
            )
        })?;
        shadow_paths.insert(file_path.clone(), shadow);
    }

    enum BatchItem {
        Group(String, Vec<usize>),
        Single(usize),
    }
    let mut items: Vec<BatchItem> = Vec::new();
    let mut group_positions: HashMap<String, usize> = HashMap::new();
    for (idx, scoped) in scoped_edits.iter().enumerate() {
        if let Some(group) = scoped.edit.group.as_ref() {
            if let Some(position) = group_positions.get(group) {
                if let BatchItem::Group(_, indices) = &mut items[*position] {
                    indices.push(idx);
                }
            } else {
                let position = items.len();
                items.push(BatchItem::Group(group.clone(), vec![idx]));
                group_positions.insert(group.clone(), position);
            }
        } else {
            items.push(BatchItem::Single(idx));
        }
    }

    let requested = scoped_edits.len();
    let mut applied = 0_usize;
    let mut skipped = 0_usize;
    let mut diagnostics = Vec::new();
    let mut lint_diagnostics = Vec::new();

    for item in items {
        let item_indices: Vec<usize> = match &item {
            BatchItem::Group(_, indices) => indices.clone(),
            BatchItem::Single(index) => vec![*index],
        };
        let mut by_file: HashMap<PathBuf, Vec<UnifiedEdit>> = HashMap::new();
        for index in &item_indices {
            let entry = &scoped_edits[*index];
            by_file
                .entry(entry.file.clone())
                .or_default()
                .push(entry.edit.clone());
        }

        let mut backups: HashMap<PathBuf, Vec<u8>> = HashMap::new();
        for file_path in by_file.keys() {
            let shadow = shadow_paths.get(file_path).ok_or_else(|| {
                anyhow::anyhow!(
                    "internal error: missing shadow file for {}",
                    file_path.display()
                )
            })?;
            backups.insert(file_path.clone(), fs::read(shadow)?);
        }

        let mut item_failed = false;
        let mut item_fail_message: Option<String> = None;

        for (file_path, file_edits) in &by_file {
            let shadow = shadow_paths.get(file_path).ok_or_else(|| {
                anyhow::anyhow!(
                    "internal error: missing shadow file for {}",
                    file_path.display()
                )
            })?;
            let mut doc = UnifiedDocument::derive(shadow, UnifiedDeriveOptions::default())?;
            if lint {
                lint_diagnostics.extend(doc.lint_edits(file_edits.as_slice()));
            }

            let report = doc.apply_edits(file_edits.as_slice())?;
            if report.skipped > 0 || !report.diagnostics.is_empty() {
                item_failed = true;
                item_fail_message = report
                    .diagnostics
                    .first()
                    .map(|diag| format!("{}: {}", diag.code, diag.message))
                    .or_else(|| Some("target id not found in derived content".to_string()));
            }
            for diag in report.diagnostics {
                diagnostics.push(serde_json::json!({
                    "severity": format!("{:?}", diag.severity),
                    "code": diag.code,
                    "message": diag.message,
                    "id": diag.id,
                }));
            }

            let save_result = doc.save_as(
                shadow,
                &ApplyOptions {
                    source_override: None,
                    force,
                },
            )?;
            if !save_result.warnings.is_empty() {
                item_failed = true;
                if item_fail_message.is_none() {
                    item_fail_message = save_result.warnings.first().cloned();
                }
            }
            diagnostics.extend(save_result.warnings.iter().map(|warning| {
                serde_json::json!({
                    "severity": "Warning",
                    "code": "apply_warning",
                    "message": warning,
                    "id": null,
                })
            }));
        }

        if item_failed {
            for (file_path, bytes) in backups {
                let shadow = shadow_paths.get(&file_path).ok_or_else(|| {
                    anyhow::anyhow!(
                        "internal error: missing shadow file for {}",
                        file_path.display()
                    )
                })?;
                fs::write(shadow, bytes)?;
            }
            skipped += item_indices.len();
            let message =
                item_fail_message.unwrap_or_else(|| "group transaction aborted".to_string());
            match &item {
                BatchItem::Group(group, _) => {
                    for index in item_indices {
                        diagnostics.push(serde_json::json!({
                            "severity": "Warning",
                            "code": "group_aborted",
                            "message": format!("edit group '{}' aborted: {}", group, message),
                            "id": scoped_edits[index].edit.id,
                        }));
                    }
                }
                BatchItem::Single(_) => {
                    diagnostics.push(serde_json::json!({
                        "severity": "Warning",
                        "code": "target_not_found",
                        "message": message,
                        "id": scoped_edits[item_indices[0]].edit.id,
                    }));
                }
            }
        } else {
            applied += item_indices.len();
        }
    }

    if unique_files.len() == 1 {
        let only_file = shadow_paths
            .keys()
            .next()
            .cloned()
            .ok_or_else(|| anyhow::anyhow!("missing staged output file"))?;
        let shadow = shadow_paths
            .get(&only_file)
            .ok_or_else(|| anyhow::anyhow!("missing staged output for file"))?;
        if in_place {
            fs::copy(shadow, &only_file)?;
        } else {
            fs::copy(shadow, &dest)?;
        }
    } else {
        for file_path in unique_files {
            let shadow = shadow_paths.get(&file_path).ok_or_else(|| {
                anyhow::anyhow!(
                    "internal error: missing shadow file for {}",
                    file_path.display()
                )
            })?;
            fs::copy(shadow, &file_path)?;
        }
    }

    let report = serde_json::json!({
        "requested": requested,
        "applied": applied,
        "skipped": skipped,
        "diagnostics": diagnostics,
        "lint_diagnostics": lint_diagnostics.iter().map(|diag| serde_json::json!({
            "severity": format!("{:?}", diag.severity),
            "code": diag.code,
            "message": diag.message,
            "id": diag.id,
        })).collect::<Vec<_>>(),
    });

    println!("{}", serde_json::to_string_pretty(&report)?);

    if strict {
        let skipped = report["skipped"].as_u64().unwrap_or(0);
        let has_diag_issues = report["diagnostics"]
            .as_array()
            .is_some_and(|items| !items.is_empty());
        if skipped > 0 || has_diag_issues {
            bail!(
                "strict mode failed: skipped={skipped}, diagnostics={}",
                report["diagnostics"]
                    .as_array()
                    .map_or(0, |items| items.len())
            );
        }
    }

    Ok(())
}

fn parse_edits_json(path: &Path, base_file: &Path) -> Result<Vec<ScopedEdit>> {
    let data = fs::read_to_string(path).map_err(|err| {
        anyhow::anyhow!("failed to read --edits-json file {}: {err}", path.display())
    })?;
    let raw: Vec<JsonEditSpec> = serde_json::from_str(data.as_str()).map_err(|err| {
        anyhow::anyhow!(
            "failed to parse --edits-json file {} as JSON array of edits: {err}",
            path.display()
        )
    })?;
    raw.into_iter()
        .map(|item| {
            let id = item.id.trim().to_string();
            if id.is_empty() {
                bail!("invalid --edits-json entry: missing id");
            }
            let mut edit = UnifiedEdit::new(id, item.text.unwrap_or_default());
            if let Some(group) = item.group {
                edit = edit.with_group(group);
            }
            if let Some(payload) = item.payload {
                let payload = match payload {
                    JsonEditPayload::XlsxCellStyle {
                        bold,
                        italic,
                        number_format,
                    } => {
                        let mut patch = CellStylePatch::new();
                        if let Some(value) = bold {
                            patch.set_bold(value);
                        }
                        if let Some(value) = italic {
                            patch.set_italic(value);
                        }
                        if let Some(value) = number_format {
                            patch.set_number_format(value);
                        }
                        UnifiedEditPayload::XlsxCellStyle(patch)
                    }
                    JsonEditPayload::PptxTextStyle {
                        bold,
                        italic,
                        font_size,
                        font_color,
                        font_name,
                    } => {
                        let mut patch = PptxTextStylePatch::new();
                        if let Some(value) = bold {
                            patch.set_bold(value);
                        }
                        if let Some(value) = italic {
                            patch.set_italic(value);
                        }
                        if let Some(value) = font_size {
                            patch.set_font_size(value);
                        }
                        if let Some(value) = font_color {
                            patch.set_font_color(value);
                        }
                        if let Some(value) = font_name {
                            patch.set_font_name(value);
                        }
                        UnifiedEditPayload::PptxTextStyle(patch)
                    }
                };
                edit = edit.with_payload(payload);
            }
            let file = item
                .file
                .map(PathBuf::from)
                .unwrap_or_else(|| base_file.to_path_buf());
            Ok(ScopedEdit { file, edit })
        })
        .collect()
}

fn parse_edit_spec(spec: &str) -> Result<UnifiedEdit> {
    let (id, text) = spec
        .split_once('=')
        .ok_or_else(|| anyhow::anyhow!("invalid --edit spec (expected <id>=<text>): {spec}"))?;
    let id = id.trim();
    if id.is_empty() {
        bail!("invalid --edit spec: missing id");
    }
    Ok(UnifiedEdit::new(id, text))
}

#[cfg(test)]
mod tests {
    use std::{
        fs,
        path::{Path, PathBuf},
    };

    use super::{parse_edit_spec, parse_edits_json};

    #[test]
    fn parse_edit_spec_happy_path() {
        let edit = parse_edit_spec("sheet:Data/cell:A1=42").expect("parse");
        assert_eq!(edit.id, "sheet:Data/cell:A1");
        assert_eq!(edit.text, "42");
    }

    #[test]
    fn parse_edit_spec_allows_equals_in_text() {
        let edit = parse_edit_spec("paragraph:2=a=b=c").expect("parse");
        assert_eq!(edit.id, "paragraph:2");
        assert_eq!(edit.text, "a=b=c");
    }

    #[test]
    fn parse_edits_json_happy_path() {
        let dir = tempfile::tempdir().expect("tempdir");
        let path = dir.path().join("edits.json");
        fs::write(
            &path,
            r#"[{"id":"sheet:Data/cell:A1","text":"100"},{"id":"paragraph:1","text":"hello"}]"#,
        )
        .expect("write json");

        let edits =
            parse_edits_json(path.as_path(), Path::new("base.xlsx")).expect("parse json edits");
        assert_eq!(edits.len(), 2);
        assert_eq!(edits[0].edit.id, "sheet:Data/cell:A1");
        assert_eq!(edits[0].edit.text, "100");
        assert_eq!(edits[1].edit.id, "paragraph:1");
        assert_eq!(edits[1].edit.text, "hello");
    }

    #[test]
    fn parse_edits_json_rejects_missing_id() {
        let dir = tempfile::tempdir().expect("tempdir");
        let path = dir.path().join("edits.json");
        fs::write(&path, r#"[{"id":"   ","text":"x"}]"#).expect("write json");

        let err = parse_edits_json(path.as_path(), Path::new("base.xlsx")).expect_err("must fail");
        assert!(err.to_string().contains("missing id"));
    }

    #[test]
    fn parse_edits_json_supports_group_and_payload() {
        let dir = tempfile::tempdir().expect("tempdir");
        let path = dir.path().join("edits.json");
        fs::write(
            &path,
            r#"[{"id":"sheet:Data/cell:A1/style","group":"txn-1","payload":{"kind":"xlsx_cell_style","bold":true,"number_format":"$#,##0"}}]"#,
        )
        .expect("write json");

        let edits =
            parse_edits_json(path.as_path(), Path::new("base.xlsx")).expect("parse json edits");
        assert_eq!(edits.len(), 1);
        assert_eq!(edits[0].edit.group.as_deref(), Some("txn-1"));
        assert!(edits[0].edit.payload.is_some());
    }

    #[test]
    fn parse_edits_json_supports_file_override() {
        let dir = tempfile::tempdir().expect("tempdir");
        let path = dir.path().join("edits.json");
        fs::write(
            &path,
            r#"[{"file":"other.xlsx","id":"sheet:Data/cell:A1","text":"1"}]"#,
        )
        .expect("write json");

        let edits =
            parse_edits_json(path.as_path(), Path::new("base.xlsx")).expect("parse json edits");
        assert_eq!(edits[0].file, PathBuf::from("other.xlsx"));
    }
}
