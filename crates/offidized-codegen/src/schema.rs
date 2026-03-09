use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};

use anyhow::{bail, Context, Result};
use serde::de::DeserializeOwned;
use serde::Deserialize;

pub const SPREADSHEETML_MAIN_URI: &str =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub const WORDPROCESSINGML_MAIN_URI: &str =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub const PRESENTATIONML_MAIN_URI: &str =
    "http://schemas.openxmlformats.org/presentationml/2006/main";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TypedElementKind {
    Leaf,
    Composite,
    TypedLeaf,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TypedElementDescriptor {
    pub class_name: String,
    pub qualified_name: String,
    pub schema_path: String,
    pub type_kind: TypedElementKind,
    pub known_attributes: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct NamespaceRecord {
    #[serde(rename = "Prefix")]
    pub prefix: String,
    #[serde(rename = "Uri")]
    pub uri: String,
    #[serde(rename = "Version", default)]
    pub version: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct TypedNamespaceRecord {
    #[serde(rename = "Prefix")]
    pub prefix: String,
    #[serde(rename = "Namespace")]
    pub namespace: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct TypedSchemaEntry {
    #[serde(rename = "Name")]
    pub name: String,
    #[serde(rename = "ClassName")]
    pub class_name: String,
    #[serde(rename = "PartClassName", default)]
    pub part_class_name: Option<String>,
}

impl TypedSchemaEntry {
    #[must_use]
    pub fn type_name(&self) -> Option<&str> {
        self.name.split_once('/').map(|(type_name, _)| type_name)
    }

    #[must_use]
    pub fn qualified_name(&self) -> Option<&str> {
        self.name
            .split_once('/')
            .and_then(|(_, qualified_name)| (!qualified_name.is_empty()).then_some(qualified_name))
    }

    #[must_use]
    pub fn element_prefix(&self) -> Option<&str> {
        let segment = self.qualified_name().or_else(|| self.type_name())?;
        segment.split_once(':').map(|(prefix, _)| prefix)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
struct SchemaRecord {
    #[serde(rename = "TargetNamespace")]
    target_namespace: String,
    #[serde(rename = "Types", default)]
    types: Vec<SchemaTypeRecord>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
struct SchemaTypeRecord {
    #[serde(rename = "Name")]
    name: String,
    #[serde(rename = "ClassName")]
    class_name: String,
    #[serde(rename = "BaseClass", default)]
    base_class: Option<String>,
    #[serde(rename = "IsLeafElement", default)]
    is_leaf_element: bool,
    #[serde(rename = "IsLeafText", default)]
    is_leaf_text: bool,
    #[serde(rename = "Attributes", default)]
    attributes: Vec<SchemaAttributeRecord>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
struct SchemaAttributeRecord {
    #[serde(rename = "QName")]
    q_name: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Default, Deserialize)]
pub struct PartPaths {
    #[serde(rename = "General", default)]
    pub general: Option<String>,
    #[serde(rename = "Excel", default)]
    pub excel: Option<String>,
    #[serde(rename = "Word", default)]
    pub word: Option<String>,
    #[serde(rename = "PowerPoint", default)]
    pub power_point: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct PartChildDefinition {
    #[serde(rename = "Name")]
    pub name: String,
    #[serde(rename = "ApiName", default)]
    pub api_name: Option<String>,
    #[serde(rename = "HasFixedContent", default)]
    pub has_fixed_content: bool,
    #[serde(rename = "IsDataPartReference", default)]
    pub is_data_part_reference: bool,
    #[serde(rename = "IsSpecialEmbeddedPart", default)]
    pub is_special_embedded_part: bool,
    #[serde(rename = "MaxOccursGreatThanOne", default)]
    pub max_occurs_great_than_one: bool,
    #[serde(rename = "MinOccursIsNonZero", default)]
    pub min_occurs_is_non_zero: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize)]
pub struct PartDefinition {
    #[serde(rename = "Name")]
    pub name: String,
    #[serde(rename = "Base")]
    pub base: String,
    #[serde(rename = "RelationshipType", default)]
    pub relationship_type: Option<String>,
    #[serde(rename = "Target", default)]
    pub target: Option<String>,
    #[serde(rename = "Root", default)]
    pub root: Option<String>,
    #[serde(rename = "RootElement", default)]
    pub root_element: Option<String>,
    #[serde(rename = "ContentType", default)]
    pub content_type: Option<String>,
    #[serde(rename = "Extension", default)]
    pub extension: Option<String>,
    #[serde(rename = "Version", default)]
    pub version: Option<String>,
    #[serde(rename = "Paths", default)]
    pub paths: PartPaths,
    #[serde(rename = "Children", default)]
    pub children: Vec<PartChildDefinition>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpenXmlSdkData {
    pub namespaces: Vec<NamespaceRecord>,
    pub typed_namespaces: Vec<TypedNamespaceRecord>,
    pub typed_schemas: BTreeMap<String, Vec<TypedSchemaEntry>>,
    pub schema_type_kinds: BTreeMap<String, BTreeMap<String, TypedElementKind>>,
    pub schema_known_attributes: BTreeMap<String, BTreeMap<String, Vec<String>>>,
    pub parts: BTreeMap<String, PartDefinition>,
}

impl OpenXmlSdkData {
    #[must_use]
    pub fn namespace_prefix_for_uri(&self, uri: &str) -> Option<&str> {
        self.namespaces
            .iter()
            .find(|namespace| namespace.uri == uri)
            .map(|namespace| namespace.prefix.as_str())
    }

    #[must_use]
    pub fn typed_entries_for_uri(&self, uri: &str) -> Option<&[TypedSchemaEntry]> {
        let typed_stem = typed_file_stem_for_uri(uri);
        self.typed_schemas.get(&typed_stem).map(Vec::as_slice)
    }

    pub fn main_namespace_entries(&self, uri: &str) -> Result<Vec<TypedSchemaEntry>> {
        let prefix = self
            .namespace_prefix_for_uri(uri)
            .with_context(|| format!("missing namespace prefix for URI: {uri}"))?;

        let entries = self
            .typed_entries_for_uri(uri)
            .with_context(|| format!("missing typed schema entries for URI: {uri}"))?;

        let filtered = entries
            .iter()
            .filter(|entry| entry.element_prefix() == Some(prefix))
            .cloned()
            .collect::<Vec<_>>();

        if filtered.is_empty() {
            bail!("no typed entries matched namespace prefix `{prefix}` for URI `{uri}`");
        }

        Ok(filtered)
    }

    pub fn main_namespace_typed_elements(&self, uri: &str) -> Result<Vec<TypedElementDescriptor>> {
        let entries = self.main_namespace_entries(uri)?;
        let type_kinds = self
            .schema_type_kinds
            .get(uri)
            .with_context(|| format!("missing schema type metadata for URI: {uri}"))?;
        let known_attributes = self
            .schema_known_attributes
            .get(uri)
            .with_context(|| format!("missing schema attribute metadata for URI: {uri}"))?;

        let mut descriptors = entries
            .into_iter()
            .filter_map(|entry| {
                let qualified_name = entry.qualified_name()?.to_owned();
                let schema_path = format!("/{qualified_name}");
                let type_kind = type_kinds
                    .get(&entry.name)
                    .copied()
                    .or_else(|| infer_type_kind_from_type_name(entry.type_name()))
                    .unwrap_or(TypedElementKind::Composite);
                let known_attributes = known_attributes
                    .get(&entry.name)
                    .or_else(|| {
                        entry
                            .qualified_name()
                            .and_then(|name| known_attributes.get(name))
                    })
                    .cloned()
                    .unwrap_or_default();

                Some(TypedElementDescriptor {
                    class_name: entry.class_name,
                    qualified_name,
                    schema_path,
                    type_kind,
                    known_attributes,
                })
            })
            .collect::<Vec<_>>();

        if descriptors.is_empty() {
            bail!("no typed element entries found for URI `{uri}`");
        }

        descriptors.sort_by(|left, right| {
            left.schema_path
                .cmp(&right.schema_path)
                .then(left.class_name.cmp(&right.class_name))
        });

        Ok(descriptors)
    }
}

pub fn load_openxml_sdk_data<P: AsRef<Path>>(data_root: P) -> Result<OpenXmlSdkData> {
    let data_root = data_root.as_ref();

    let namespaces_path = data_root.join("namespaces.json");
    let namespaces = read_json_file::<Vec<NamespaceRecord>>(&namespaces_path)
        .with_context(|| format!("failed to parse {}", display_path(&namespaces_path)))?;

    let typed_root = data_root.join("typed");
    let (typed_namespaces, typed_schemas) = load_typed_data(&typed_root)?;

    let schemas_root = data_root.join("schemas");
    let schema_type_kinds = load_schema_type_kinds(&schemas_root)?;
    let schema_known_attributes = load_schema_known_attributes(&schemas_root)?;

    let parts_root = data_root.join("parts");
    let parts = load_parts(&parts_root)?;

    Ok(OpenXmlSdkData {
        namespaces,
        typed_namespaces,
        typed_schemas,
        schema_type_kinds,
        schema_known_attributes,
        parts,
    })
}

#[must_use]
pub fn typed_file_stem_for_uri(uri: &str) -> String {
    let without_scheme = uri
        .strip_prefix("http://")
        .or_else(|| uri.strip_prefix("https://"))
        .unwrap_or(uri);

    let mut normalized = String::with_capacity(without_scheme.len());
    for ch in without_scheme.chars() {
        if ch.is_ascii_alphanumeric() {
            normalized.push(ch);
        } else {
            normalized.push('_');
        }
    }

    normalized
}

fn load_typed_data(
    typed_root: &Path,
) -> Result<(
    Vec<TypedNamespaceRecord>,
    BTreeMap<String, Vec<TypedSchemaEntry>>,
)> {
    let typed_files = collect_json_files(typed_root)?;
    let mut typed_namespaces = None;
    let mut typed_schemas = BTreeMap::new();

    for typed_file in typed_files {
        let file_name = typed_file
            .file_name()
            .and_then(|name| name.to_str())
            .with_context(|| format!("invalid UTF-8 file name: {}", display_path(&typed_file)))?;

        if file_name == "namespaces.json" {
            let records = read_json_file::<Vec<TypedNamespaceRecord>>(&typed_file)
                .with_context(|| format!("failed to parse {}", display_path(&typed_file)))?;
            typed_namespaces = Some(records);
            continue;
        }

        let entries = read_json_file::<Vec<TypedSchemaEntry>>(&typed_file)
            .with_context(|| format!("failed to parse {}", display_path(&typed_file)))?;
        let stem = typed_file
            .file_stem()
            .and_then(|name| name.to_str())
            .with_context(|| format!("invalid UTF-8 file stem: {}", display_path(&typed_file)))?;

        if typed_schemas.insert(stem.to_owned(), entries).is_some() {
            bail!("duplicate typed schema file stem: {stem}");
        }
    }

    let typed_namespaces =
        typed_namespaces.context("missing required typed/namespaces.json file")?;

    Ok((typed_namespaces, typed_schemas))
}

fn load_parts(parts_root: &Path) -> Result<BTreeMap<String, PartDefinition>> {
    let part_files = collect_json_files(parts_root)?;
    let mut parts = BTreeMap::new();

    for part_file in part_files {
        let part = read_json_file::<PartDefinition>(&part_file)
            .with_context(|| format!("failed to parse {}", display_path(&part_file)))?;

        if parts.insert(part.name.clone(), part).is_some() {
            bail!("duplicate part name in parts registry");
        }
    }

    Ok(parts)
}

fn load_schema_type_kinds(
    schemas_root: &Path,
) -> Result<BTreeMap<String, BTreeMap<String, TypedElementKind>>> {
    let schema_files = collect_json_files(schemas_root)?;
    let mut schemas = BTreeMap::new();

    for schema_file in schema_files {
        let schema = read_json_file::<SchemaRecord>(&schema_file)
            .with_context(|| format!("failed to parse {}", display_path(&schema_file)))?;
        let kind_by_name = derive_schema_type_kinds(&schema.types);

        if schemas
            .insert(schema.target_namespace.clone(), kind_by_name)
            .is_some()
        {
            bail!(
                "duplicate schema namespace metadata for URI: {}",
                schema.target_namespace
            );
        }
    }

    Ok(schemas)
}

fn load_schema_known_attributes(
    schemas_root: &Path,
) -> Result<BTreeMap<String, BTreeMap<String, Vec<String>>>> {
    let schema_files = collect_json_files(schemas_root)?;
    let mut schemas = BTreeMap::new();

    for schema_file in schema_files {
        let schema = read_json_file::<SchemaRecord>(&schema_file)
            .with_context(|| format!("failed to parse {}", display_path(&schema_file)))?;
        let attributes_by_name = derive_schema_known_attributes(&schema.types);

        if schemas
            .insert(schema.target_namespace.clone(), attributes_by_name)
            .is_some()
        {
            bail!(
                "duplicate schema namespace attribute metadata for URI: {}",
                schema.target_namespace
            );
        }
    }

    Ok(schemas)
}

fn derive_schema_type_kinds(types: &[SchemaTypeRecord]) -> BTreeMap<String, TypedElementKind> {
    let mut type_by_class_name = BTreeMap::new();
    for schema_type in types {
        type_by_class_name
            .entry(schema_type.class_name.clone())
            .or_insert_with(|| schema_type.clone());
    }

    let mut memo = BTreeMap::new();
    let mut kind_by_name = BTreeMap::new();
    for schema_type in types {
        if let Some(kind) = resolve_type_kind(
            &schema_type.class_name,
            &type_by_class_name,
            &mut memo,
            &mut Vec::new(),
        ) {
            kind_by_name.insert(schema_type.name.clone(), kind);
        }
    }

    kind_by_name
}

fn derive_schema_known_attributes(types: &[SchemaTypeRecord]) -> BTreeMap<String, Vec<String>> {
    let mut type_by_class_name = BTreeMap::new();
    for schema_type in types {
        type_by_class_name
            .entry(schema_type.class_name.clone())
            .or_insert_with(|| schema_type.clone());
    }

    let mut resolved_by_class_name = BTreeMap::new();
    let mut known_attributes = BTreeMap::<String, BTreeSet<String>>::new();

    for schema_type in types {
        let attributes = resolve_known_attributes(
            &schema_type.class_name,
            &type_by_class_name,
            &mut resolved_by_class_name,
            &mut Vec::new(),
        );

        merge_known_attributes(
            &mut known_attributes,
            schema_type.name.clone(),
            attributes.clone(),
        );

        if let Some((_, qualified_name)) = schema_type.name.split_once('/') {
            merge_known_attributes(
                &mut known_attributes,
                qualified_name.to_owned(),
                attributes.clone(),
            );
        }
    }

    known_attributes
        .into_iter()
        .map(|(name, attributes)| (name, attributes.into_iter().collect()))
        .collect()
}

fn resolve_known_attributes(
    class_name: &str,
    type_by_class_name: &BTreeMap<String, SchemaTypeRecord>,
    resolved_by_class_name: &mut BTreeMap<String, BTreeSet<String>>,
    visiting: &mut Vec<String>,
) -> BTreeSet<String> {
    if let Some(cached) = resolved_by_class_name.get(class_name) {
        return cached.clone();
    }

    if visiting.iter().any(|current| current == class_name) {
        return BTreeSet::new();
    }

    let Some(schema_type) = type_by_class_name.get(class_name) else {
        return BTreeSet::new();
    };
    visiting.push(class_name.to_owned());

    let mut attributes = schema_type
        .attributes
        .iter()
        .filter_map(|attribute| normalize_attribute_q_name(&attribute.q_name))
        .collect::<BTreeSet<_>>();

    if let Some(base_class) = schema_type.base_class.as_deref() {
        attributes.extend(resolve_known_attributes(
            base_class,
            type_by_class_name,
            resolved_by_class_name,
            visiting,
        ));
    }

    let _ = visiting.pop();
    resolved_by_class_name.insert(class_name.to_owned(), attributes.clone());
    attributes
}

fn normalize_attribute_q_name(q_name: &str) -> Option<String> {
    let normalized = q_name.strip_prefix(':').unwrap_or(q_name).trim();
    (!normalized.is_empty()).then_some(normalized.to_owned())
}

fn merge_known_attributes(
    known_attributes: &mut BTreeMap<String, BTreeSet<String>>,
    key: String,
    attrs: BTreeSet<String>,
) {
    known_attributes
        .entry(key)
        .and_modify(|existing| existing.extend(attrs.iter().cloned()))
        .or_insert(attrs);
}

fn resolve_type_kind(
    class_name: &str,
    type_by_class_name: &BTreeMap<String, SchemaTypeRecord>,
    memo: &mut BTreeMap<String, Option<TypedElementKind>>,
    visiting: &mut Vec<String>,
) -> Option<TypedElementKind> {
    if let Some(cached) = memo.get(class_name) {
        return *cached;
    }

    if visiting.iter().any(|current| current == class_name) {
        return None;
    }

    let schema_type = type_by_class_name.get(class_name)?;
    visiting.push(class_name.to_owned());

    let kind = direct_type_kind(schema_type).or_else(|| {
        schema_type.base_class.as_deref().and_then(|base_class| {
            resolve_type_kind(base_class, type_by_class_name, memo, visiting)
        })
    });

    let _ = visiting.pop();
    memo.insert(class_name.to_owned(), kind);
    kind
}

fn direct_type_kind(schema_type: &SchemaTypeRecord) -> Option<TypedElementKind> {
    if schema_type.is_leaf_text
        || schema_type.base_class.as_deref() == Some("OpenXmlLeafTextElement")
    {
        return Some(TypedElementKind::TypedLeaf);
    }

    if schema_type.is_leaf_element
        || schema_type.base_class.as_deref() == Some("OpenXmlLeafElement")
    {
        return Some(TypedElementKind::Leaf);
    }

    if schema_type.base_class.as_deref() == Some("OpenXmlCompositeElement") {
        return Some(TypedElementKind::Composite);
    }

    None
}

fn infer_type_kind_from_type_name(type_name: Option<&str>) -> Option<TypedElementKind> {
    let type_name = type_name?;

    if type_name.starts_with("xsd:") || type_name.starts_with("xs:") || type_name.contains(":ST_") {
        return Some(TypedElementKind::TypedLeaf);
    }

    if type_name.contains(":CT_") {
        return Some(TypedElementKind::Composite);
    }

    None
}

fn collect_json_files(root: &Path) -> Result<Vec<PathBuf>> {
    let mut files = fs::read_dir(root)
        .with_context(|| format!("failed to read directory {}", display_path(root)))?
        .collect::<std::result::Result<Vec<_>, _>>()
        .with_context(|| format!("failed to enumerate directory {}", display_path(root)))?
        .into_iter()
        .map(|entry| entry.path())
        .filter(|path| path.extension().and_then(|ext| ext.to_str()) == Some("json"))
        .collect::<Vec<_>>();

    files.sort_unstable();
    Ok(files)
}

fn read_json_file<T: DeserializeOwned>(path: &Path) -> Result<T> {
    let bytes = fs::read(path).with_context(|| format!("failed to read {}", display_path(path)))?;
    serde_json::from_slice::<T>(&bytes)
        .with_context(|| format!("failed to deserialize {}", display_path(path)))
}

fn display_path(path: &Path) -> String {
    path.to_string_lossy().into_owned()
}

#[cfg(test)]
#[allow(clippy::panic_in_result_fn)]
mod tests {
    use std::path::PathBuf;

    use anyhow::Result;

    use super::{
        load_openxml_sdk_data, typed_file_stem_for_uri, OpenXmlSdkData, TypedElementKind,
        TypedSchemaEntry, PRESENTATIONML_MAIN_URI, SPREADSHEETML_MAIN_URI,
        WORDPROCESSINGML_MAIN_URI,
    };

    #[test]
    fn load_openxml_sdk_data_reads_expected_sections() -> Result<()> {
        let Some(data_root) = test_data_root_if_present() else {
            eprintln!("SKIP: Open-XML-SDK reference data not found");
            return Ok(());
        };
        let data = load_openxml_sdk_data(data_root)?;

        assert!(!data.namespaces.is_empty());
        assert!(!data.typed_namespaces.is_empty());
        assert!(!data.typed_schemas.is_empty());
        assert!(!data.schema_type_kinds.is_empty());
        assert!(!data.schema_known_attributes.is_empty());
        assert!(!data.parts.is_empty());

        assert_eq!(
            data.namespace_prefix_for_uri(SPREADSHEETML_MAIN_URI),
            Some("x")
        );
        assert_eq!(
            data.namespace_prefix_for_uri(WORDPROCESSINGML_MAIN_URI),
            Some("w")
        );
        assert_eq!(
            data.namespace_prefix_for_uri(PRESENTATIONML_MAIN_URI),
            Some("p")
        );

        assert!(data
            .typed_schemas
            .contains_key(&typed_file_stem_for_uri(SPREADSHEETML_MAIN_URI)));
        assert!(data
            .typed_schemas
            .contains_key(&typed_file_stem_for_uri(WORDPROCESSINGML_MAIN_URI)));
        assert!(data
            .typed_schemas
            .contains_key(&typed_file_stem_for_uri(PRESENTATIONML_MAIN_URI)));
        assert!(data.schema_type_kinds.contains_key(SPREADSHEETML_MAIN_URI));
        assert!(data
            .schema_type_kinds
            .contains_key(WORDPROCESSINGML_MAIN_URI));
        assert!(data.schema_type_kinds.contains_key(PRESENTATIONML_MAIN_URI));
        assert!(data
            .schema_known_attributes
            .contains_key(SPREADSHEETML_MAIN_URI));
        assert!(data
            .schema_known_attributes
            .contains_key(WORDPROCESSINGML_MAIN_URI));
        assert!(data
            .schema_known_attributes
            .contains_key(PRESENTATIONML_MAIN_URI));

        assert!(data.parts.contains_key("WorkbookPart"));
        assert!(data.parts.contains_key("MainDocumentPart"));
        assert!(data.parts.contains_key("PresentationPart"));

        Ok(())
    }

    #[test]
    fn typed_schema_entry_extracts_prefix_from_composite_name() {
        let entry = TypedSchemaEntry {
            name: "xsd:string/p:attrName".to_string(),
            class_name: "AttributeName".to_string(),
            part_class_name: None,
        };
        assert_eq!(entry.type_name(), Some("xsd:string"));
        assert_eq!(entry.qualified_name(), Some("p:attrName"));
        assert_eq!(entry.element_prefix(), Some("p"));

        let no_colon = TypedSchemaEntry {
            name: "plainIdentifier".to_string(),
            class_name: "Plain".to_string(),
            part_class_name: None,
        };
        assert_eq!(no_colon.type_name(), None);
        assert_eq!(no_colon.qualified_name(), None);
        assert_eq!(no_colon.element_prefix(), None);
    }

    #[test]
    fn main_namespace_entries_are_filtered_by_prefix() -> Result<()> {
        let Some(data_root) = test_data_root_if_present() else {
            eprintln!("SKIP: Open-XML-SDK reference data not found");
            return Ok(());
        };
        let data: OpenXmlSdkData = load_openxml_sdk_data(data_root)?;

        let all_presentation_entries = data
            .typed_entries_for_uri(PRESENTATIONML_MAIN_URI)
            .map_or(0, |entries| entries.len());
        let main_entries = data.main_namespace_entries(PRESENTATIONML_MAIN_URI)?;

        assert!(!main_entries.is_empty());
        assert!(main_entries
            .iter()
            .all(|entry| entry.element_prefix() == Some("p")));
        assert!(all_presentation_entries > main_entries.len());

        Ok(())
    }

    #[test]
    fn main_namespace_typed_elements_include_kind_and_paths() -> Result<()> {
        let Some(data_root) = test_data_root_if_present() else {
            eprintln!("SKIP: Open-XML-SDK reference data not found");
            return Ok(());
        };
        let data: OpenXmlSdkData = load_openxml_sdk_data(data_root)?;

        let spreadsheet = data.main_namespace_typed_elements(SPREADSHEETML_MAIN_URI)?;
        let workbook = spreadsheet
            .iter()
            .find(|entry| entry.class_name == "Workbook")
            .expect("Workbook typed descriptor should exist");
        assert_eq!(workbook.qualified_name, "x:workbook");
        assert_eq!(workbook.schema_path, "/x:workbook");
        assert_eq!(workbook.type_kind, TypedElementKind::Composite);

        let word = data.main_namespace_typed_elements(WORDPROCESSINGML_MAIN_URI)?;
        let text = word
            .iter()
            .find(|entry| entry.class_name == "Text")
            .expect("Wordprocessing Text typed descriptor should exist");
        assert_eq!(text.qualified_name, "w:t");
        assert_eq!(text.type_kind, TypedElementKind::TypedLeaf);
        assert!(text.known_attributes.contains(&"xml:space".to_string()));

        let presentation = data.main_namespace_typed_elements(PRESENTATIONML_MAIN_URI)?;
        let slide_all = presentation
            .iter()
            .find(|entry| entry.class_name == "SlideAll")
            .expect("SlideAll typed descriptor should exist");
        assert_eq!(slide_all.qualified_name, "p:sldAll");
        assert_eq!(slide_all.type_kind, TypedElementKind::Leaf);

        Ok(())
    }

    fn test_data_root_if_present() -> Option<PathBuf> {
        let root =
            PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../references/Open-XML-SDK/data");
        root.exists().then_some(root)
    }
}
