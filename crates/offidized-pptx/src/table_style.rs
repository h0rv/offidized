//! Table style support for PowerPoint presentations.
//!
//! This module provides GUID-based table style identification and table style options
//! for controlling which parts of a table receive special formatting (header row,
//! total row, banded rows/columns, etc.).

use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

/// Built-in PowerPoint table style types.
///
/// PowerPoint provides a set of pre-defined table styles categorized by theme
/// (Light, Medium, Dark) and accent color combinations.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TableStyleType {
    NoStyleNoGrid,
    NoStyleTableGrid,
    ThemedStyle1Accent1,
    ThemedStyle1Accent2,
    ThemedStyle1Accent3,
    ThemedStyle1Accent4,
    ThemedStyle1Accent5,
    ThemedStyle1Accent6,
    ThemedStyle2Accent1,
    ThemedStyle2Accent2,
    ThemedStyle2Accent3,
    ThemedStyle2Accent4,
    ThemedStyle2Accent5,
    ThemedStyle2Accent6,
    LightStyle1,
    LightStyle1Accent1,
    LightStyle1Accent2,
    LightStyle1Accent3,
    LightStyle1Accent4,
    LightStyle1Accent5,
    LightStyle1Accent6,
    LightStyle2,
    LightStyle2Accent1,
    LightStyle2Accent2,
    LightStyle2Accent3,
    LightStyle2Accent4,
    LightStyle2Accent5,
    LightStyle2Accent6,
    LightStyle3,
    LightStyle3Accent1,
    LightStyle3Accent2,
    LightStyle3Accent3,
    LightStyle3Accent4,
    LightStyle3Accent5,
    LightStyle3Accent6,
    MediumStyle1,
    MediumStyle1Accent1,
    MediumStyle1Accent2,
    MediumStyle1Accent3,
    MediumStyle1Accent4,
    MediumStyle1Accent5,
    MediumStyle1Accent6,
    MediumStyle2,
    MediumStyle2Accent1,
    MediumStyle2Accent2,
    MediumStyle2Accent3,
    MediumStyle2Accent4,
    MediumStyle2Accent5,
    MediumStyle2Accent6,
    MediumStyle3,
    MediumStyle3Accent1,
    MediumStyle3Accent2,
    MediumStyle3Accent3,
    MediumStyle3Accent4,
    MediumStyle3Accent5,
    MediumStyle3Accent6,
    MediumStyle4,
    MediumStyle4Accent1,
    MediumStyle4Accent2,
    MediumStyle4Accent3,
    MediumStyle4Accent4,
    MediumStyle4Accent5,
    MediumStyle4Accent6,
    DarkStyle1,
    DarkStyle1Accent1,
    DarkStyle1Accent2,
    DarkStyle1Accent3,
    DarkStyle1Accent4,
    DarkStyle1Accent5,
    DarkStyle1Accent6,
    DarkStyle2,
    DarkStyle2Accent1Accent2,
    DarkStyle2Accent3Accent4,
    DarkStyle2Accent5Accent6,
}

/// A PowerPoint table style identified by GUID.
///
/// PowerPoint table styles are referenced by a unique GUID string. This struct
/// holds both the GUID and the human-readable name for the style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableStyle {
    /// Unique identifier for this table style.
    guid: String,
    /// Human-readable name of the style.
    name: String,
}

impl TableStyle {
    /// Creates a new table style with the specified GUID and name.
    pub fn new(guid: impl Into<String>, name: impl Into<String>) -> Self {
        Self {
            guid: guid.into(),
            name: name.into(),
        }
    }

    /// Returns the GUID of this table style.
    pub fn guid(&self) -> &str {
        &self.guid
    }

    /// Returns the name of this table style.
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// Table style options control which parts of a table receive special formatting.
///
/// These boolean flags correspond to PowerPoint's table style options that determine
/// whether header rows, total rows, banded rows/columns, and first/last columns
/// receive special styling treatment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableStyleOptions {
    /// Apply special formatting to the first row (header row).
    pub first_row: bool,
    /// Apply special formatting to the last row (total row).
    pub last_row: bool,
    /// Apply special formatting to the first column.
    pub first_col: bool,
    /// Apply special formatting to the last column.
    pub last_col: bool,
    /// Apply alternating formatting to rows (banded rows).
    pub banded_rows: bool,
    /// Apply alternating formatting to columns (banded columns).
    pub banded_cols: bool,
    /// Unknown attributes on `<a:tblPr>` preserved for roundtrip fidelity.
    pub unknown_attrs: Vec<(String, String)>,
}

impl TableStyleOptions {
    /// Creates a new `TableStyleOptions` with all options set to `false`.
    pub fn new() -> Self {
        Self {
            first_row: false,
            last_row: false,
            first_col: false,
            last_col: false,
            banded_rows: false,
            banded_cols: false,
            unknown_attrs: Vec::new(),
        }
    }

    /// Creates a `TableStyleOptions` with default PowerPoint settings (first row enabled).
    pub fn default_pptx() -> Self {
        Self {
            first_row: true,
            last_row: false,
            first_col: false,
            last_col: false,
            banded_rows: false,
            banded_cols: false,
            unknown_attrs: Vec::new(),
        }
    }
}

impl Default for TableStyleOptions {
    fn default() -> Self {
        Self::new()
    }
}

/// Returns the built-in table style for the specified style type.
///
/// PowerPoint comes with a set of pre-defined table styles identified by GUIDs.
/// This function returns the `TableStyle` for any of the built-in style types.
///
/// # Example
///
/// ```
/// use offidized_pptx::table_style::{get_builtin_style, TableStyleType};
///
/// let style = get_builtin_style(TableStyleType::MediumStyle2Accent3);
/// assert_eq!(style.name(), "Medium Style 2 - Accent 3");
/// assert_eq!(style.guid(), "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}");
/// ```
pub fn get_builtin_style(style_type: TableStyleType) -> TableStyle {
    match style_type {
        TableStyleType::NoStyleNoGrid => TableStyle::new(
            "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
            "No Style, No Grid",
        ),
        TableStyleType::NoStyleTableGrid => TableStyle::new(
            "{5940675A-B579-460E-94D1-54222C63F5DA}",
            "No Style, Table Grid",
        ),
        TableStyleType::ThemedStyle1Accent1 => TableStyle::new(
            "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}",
            "Themed Style 1 - Accent 1",
        ),
        TableStyleType::ThemedStyle1Accent2 => TableStyle::new(
            "{284E427A-3D55-4303-BF80-6455036E1DE7}",
            "Themed Style 1 - Accent 2",
        ),
        TableStyleType::ThemedStyle1Accent3 => TableStyle::new(
            "{69C7853C-536D-4A76-A0AE-DD22124D55A5}",
            "Themed Style 1 - Accent 3",
        ),
        TableStyleType::ThemedStyle1Accent4 => TableStyle::new(
            "{775DCB02-9BB8-47FD-8907-85C794F793BA}",
            "Themed Style 1 - Accent 4",
        ),
        TableStyleType::ThemedStyle1Accent5 => TableStyle::new(
            "{35758FB7-9AC5-4552-8A53-C91805E547FA}",
            "Themed Style 1 - Accent 5",
        ),
        TableStyleType::ThemedStyle1Accent6 => TableStyle::new(
            "{08FB837D-C827-4EFA-A057-4D05807E0F7C}",
            "Themed Style 1 - Accent 6",
        ),
        TableStyleType::ThemedStyle2Accent1 => TableStyle::new(
            "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}",
            "Themed Style 2 - Accent 1",
        ),
        TableStyleType::ThemedStyle2Accent2 => TableStyle::new(
            "{18603FDC-E32A-AB5-989C-0864C3EAD2B8}",
            "Themed Style 2 - Accent 2",
        ),
        TableStyleType::ThemedStyle2Accent3 => TableStyle::new(
            "{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}",
            "Themed Style 2 - Accent 3",
        ),
        TableStyleType::ThemedStyle2Accent4 => TableStyle::new(
            "{E269D01E-BC32-4049-B463-5C60D7B0CCD2}",
            "Themed Style 2 - Accent 4",
        ),
        TableStyleType::ThemedStyle2Accent5 => TableStyle::new(
            "{327F97BB-C833-4FB7-BDE5-3F7075034690}",
            "Themed Style 2 - Accent 5",
        ),
        TableStyleType::ThemedStyle2Accent6 => TableStyle::new(
            "{638B1855-1B75-4FBE-930C-398BA8C253C6}",
            "Themed Style 2 - Accent 6",
        ),
        TableStyleType::LightStyle1 => {
            TableStyle::new("{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}", "Light Style 1")
        }
        TableStyleType::LightStyle1Accent1 => TableStyle::new(
            "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
            "Light Style 1 - Accent 1",
        ),
        TableStyleType::LightStyle1Accent2 => TableStyle::new(
            "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}",
            "Light Style 1 - Accent 2",
        ),
        TableStyleType::LightStyle1Accent3 => TableStyle::new(
            "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}",
            "Light Style 1 - Accent 3",
        ),
        TableStyleType::LightStyle1Accent4 => TableStyle::new(
            "{D27102A9-8310-4765-A935-A1911B00CA55}",
            "Light Style 1 - Accent 4",
        ),
        TableStyleType::LightStyle1Accent5 => TableStyle::new(
            "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}",
            "Light Style 1 - Accent 5",
        ),
        TableStyleType::LightStyle1Accent6 => TableStyle::new(
            "{68D230F3-CF80-4859-8CE7-A43EE81993B5}",
            "Light Style 1 - Accent 6",
        ),
        TableStyleType::LightStyle2 => {
            TableStyle::new("{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}", "Light Style 2")
        }
        TableStyleType::LightStyle2Accent1 => TableStyle::new(
            "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}",
            "Light Style 2 - Accent 1",
        ),
        TableStyleType::LightStyle2Accent2 => TableStyle::new(
            "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}",
            "Light Style 2 - Accent 2",
        ),
        TableStyleType::LightStyle2Accent3 => TableStyle::new(
            "{F2DE63D5-997A-4646-A377-4702673A728D}",
            "Light Style 2 - Accent 3",
        ),
        TableStyleType::LightStyle2Accent4 => TableStyle::new(
            "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}",
            "Light Style 2 - Accent 4",
        ),
        TableStyleType::LightStyle2Accent5 => TableStyle::new(
            "{5A111915-BE36-4E01-A7E5-04B1672EAD32}",
            "Light Style 2 - Accent 5",
        ),
        TableStyleType::LightStyle2Accent6 => TableStyle::new(
            "{912C8C85-51F0-491E-9774-3900AFEF0FD7}",
            "Light Style 2 - Accent 6",
        ),
        TableStyleType::LightStyle3 => {
            TableStyle::new("{616DA210-FB5B-4158-B5E0-FEB733F419BA}", "Light Style 3")
        }
        TableStyleType::LightStyle3Accent1 => TableStyle::new(
            "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}",
            "Light Style 3 - Accent 1",
        ),
        TableStyleType::LightStyle3Accent2 => TableStyle::new(
            "{5DA37D80-6434-44D0-A028-1B22A696006F}",
            "Light Style 3 - Accent 2",
        ),
        TableStyleType::LightStyle3Accent3 => TableStyle::new(
            "{8799B23B-EC83-4686-B30A-512413B5E67A}",
            "Light Style 3 - Accent 3",
        ),
        TableStyleType::LightStyle3Accent4 => TableStyle::new(
            "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}",
            "Light Style 3 - Accent 4",
        ),
        TableStyleType::LightStyle3Accent5 => TableStyle::new(
            "{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}",
            "Light Style 3 - Accent 5",
        ),
        TableStyleType::LightStyle3Accent6 => TableStyle::new(
            "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}",
            "Light Style 3 - Accent 6",
        ),
        TableStyleType::MediumStyle1 => {
            TableStyle::new("{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}", "Medium Style 1")
        }
        TableStyleType::MediumStyle1Accent1 => TableStyle::new(
            "{B301B821-A1FF-4177-AEE7-76D212191A09}",
            "Medium Style 1 - Accent 1",
        ),
        TableStyleType::MediumStyle1Accent2 => TableStyle::new(
            "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}",
            "Medium Style 1 - Accent 2",
        ),
        TableStyleType::MediumStyle1Accent3 => TableStyle::new(
            "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}",
            "Medium Style 1 - Accent 3",
        ),
        TableStyleType::MediumStyle1Accent4 => TableStyle::new(
            "{1E171933-4619-4E11-9A3F-F7608DF75F80}",
            "Medium Style 1 - Accent 4",
        ),
        TableStyleType::MediumStyle1Accent5 => TableStyle::new(
            "{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}",
            "Medium Style 1 - Accent 5",
        ),
        TableStyleType::MediumStyle1Accent6 => TableStyle::new(
            "{10A1B5D5-9B99-4C35-A422-299274C87663}",
            "Medium Style 1 - Accent 6",
        ),
        TableStyleType::MediumStyle2 => {
            TableStyle::new("{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}", "Medium Style 2")
        }
        TableStyleType::MediumStyle2Accent1 => TableStyle::new(
            "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
            "Medium Style 2 - Accent 1",
        ),
        TableStyleType::MediumStyle2Accent2 => TableStyle::new(
            "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}",
            "Medium Style 2 - Accent 2",
        ),
        TableStyleType::MediumStyle2Accent3 => TableStyle::new(
            "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
            "Medium Style 2 - Accent 3",
        ),
        TableStyleType::MediumStyle2Accent4 => TableStyle::new(
            "{00A15C55-8517-42AA-B614-E9B94910E393}",
            "Medium Style 2 - Accent 4",
        ),
        TableStyleType::MediumStyle2Accent5 => TableStyle::new(
            "{7DF18680-E054-41AD-8BC1-D1AEF772440D}",
            "Medium Style 2 - Accent 5",
        ),
        TableStyleType::MediumStyle2Accent6 => TableStyle::new(
            "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}",
            "Medium Style 2 - Accent 6",
        ),
        TableStyleType::MediumStyle3 => {
            TableStyle::new("{8EC20E35-A176-4012-BC5E-935CFFF8708E}", "Medium Style 3")
        }
        TableStyleType::MediumStyle3Accent1 => TableStyle::new(
            "{6E25E649-3F16-4E02-A733-19D2CDBF48F0}",
            "Medium Style 3 - Accent 1",
        ),
        TableStyleType::MediumStyle3Accent2 => TableStyle::new(
            "{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}",
            "Medium Style 3 - Accent 2",
        ),
        TableStyleType::MediumStyle3Accent3 => TableStyle::new(
            "{EB344D84-9AFB-497E-A393-DC336BA19D2E}",
            "Medium Style 3 - Accent 3",
        ),
        TableStyleType::MediumStyle3Accent4 => TableStyle::new(
            "{EB9631B5-78F2-41C9-869B-9F39066F8104}",
            "Medium Style 3 - Accent 4",
        ),
        TableStyleType::MediumStyle3Accent5 => TableStyle::new(
            "{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}",
            "Medium Style 3 - Accent 5",
        ),
        TableStyleType::MediumStyle3Accent6 => TableStyle::new(
            "{2A488322-F2BA-4B5B-9748-0D474271808F}",
            "Medium Style 3 - Accent 6",
        ),
        TableStyleType::MediumStyle4 => {
            TableStyle::new("{D7AC3CCA-C797-4891-BE02-D94E43425B78}", "Medium Style 4")
        }
        TableStyleType::MediumStyle4Accent1 => TableStyle::new(
            "{69CF1AB2-1976-4502-BF36-3FF5EA218861}",
            "Medium Style 4 - Accent 1",
        ),
        TableStyleType::MediumStyle4Accent2 => TableStyle::new(
            "{8A107856-5554-42FB-B03E-39F5DBC370BA}",
            "Medium Style 4 - Accent 2",
        ),
        TableStyleType::MediumStyle4Accent3 => TableStyle::new(
            "{0505E3EF-67EA-436B-97B2-0124C06EBD24}",
            "Medium Style 4 - Accent 3",
        ),
        TableStyleType::MediumStyle4Accent4 => TableStyle::new(
            "{C4B1156A-380E-4F78-BDF5-A606A8083BF9}",
            "Medium Style 4 - Accent 4",
        ),
        TableStyleType::MediumStyle4Accent5 => TableStyle::new(
            "{22838BEF-8BB2-4498-84A7-C5851F593DF1}",
            "Medium Style 4 - Accent 5",
        ),
        TableStyleType::MediumStyle4Accent6 => TableStyle::new(
            "{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}",
            "Medium Style 4 - Accent 6",
        ),
        TableStyleType::DarkStyle1 => {
            TableStyle::new("{E8034E78-7F5D-4C2E-B375-FC64B27BC917}", "Dark Style 1")
        }
        TableStyleType::DarkStyle1Accent1 => TableStyle::new(
            "{125E5076-3810-47DD-B79F-674D7AD40C01}",
            "Dark Style 1 - Accent 1",
        ),
        TableStyleType::DarkStyle1Accent2 => TableStyle::new(
            "{37CE84F3-28C3-443E-9E96-99CF82512B78}",
            "Dark Style 1 - Accent 2",
        ),
        TableStyleType::DarkStyle1Accent3 => TableStyle::new(
            "{D03447BB-5D67-496B-8E87-E561075AD55C}",
            "Dark Style 1 - Accent 3",
        ),
        TableStyleType::DarkStyle1Accent4 => TableStyle::new(
            "{E929F9F4-4A8F-4326-A1B4-22849713DDAB}",
            "Dark Style 1 - Accent 4",
        ),
        TableStyleType::DarkStyle1Accent5 => TableStyle::new(
            "{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}",
            "Dark Style 1 - Accent 5",
        ),
        TableStyleType::DarkStyle1Accent6 => TableStyle::new(
            "{AF606853-7671-496A-8E4F-DF71F8EC918B}",
            "Dark Style 1 - Accent 6",
        ),
        TableStyleType::DarkStyle2 => {
            TableStyle::new("{5202B0CA-FC54-4496-8BCA-5EF66A818D29}", "Dark Style 2")
        }
        TableStyleType::DarkStyle2Accent1Accent2 => TableStyle::new(
            "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}",
            "Dark Style 2 - Accent 1, Accent 2",
        ),
        TableStyleType::DarkStyle2Accent3Accent4 => TableStyle::new(
            "{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}",
            "Dark Style 2 - Accent 3, Accent 4",
        ),
        TableStyleType::DarkStyle2Accent5Accent6 => TableStyle::new(
            "{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}",
            "Dark Style 2 - Accent 5, Accent 6",
        ),
    }
}

/// Looks up a built-in table style by GUID.
///
/// Returns `Some(TableStyle)` if the GUID matches one of PowerPoint's built-in
/// table styles, or `None` if the GUID is not recognized.
///
/// The comparison is case-insensitive.
pub fn get_style_by_guid(guid: &str) -> Option<TableStyle> {
    // Try all built-in styles
    let all_types = [
        TableStyleType::NoStyleNoGrid,
        TableStyleType::NoStyleTableGrid,
        TableStyleType::ThemedStyle1Accent1,
        TableStyleType::ThemedStyle1Accent2,
        TableStyleType::ThemedStyle1Accent3,
        TableStyleType::ThemedStyle1Accent4,
        TableStyleType::ThemedStyle1Accent5,
        TableStyleType::ThemedStyle1Accent6,
        TableStyleType::ThemedStyle2Accent1,
        TableStyleType::ThemedStyle2Accent2,
        TableStyleType::ThemedStyle2Accent3,
        TableStyleType::ThemedStyle2Accent4,
        TableStyleType::ThemedStyle2Accent5,
        TableStyleType::ThemedStyle2Accent6,
        TableStyleType::LightStyle1,
        TableStyleType::LightStyle1Accent1,
        TableStyleType::LightStyle1Accent2,
        TableStyleType::LightStyle1Accent3,
        TableStyleType::LightStyle1Accent4,
        TableStyleType::LightStyle1Accent5,
        TableStyleType::LightStyle1Accent6,
        TableStyleType::LightStyle2,
        TableStyleType::LightStyle2Accent1,
        TableStyleType::LightStyle2Accent2,
        TableStyleType::LightStyle2Accent3,
        TableStyleType::LightStyle2Accent4,
        TableStyleType::LightStyle2Accent5,
        TableStyleType::LightStyle2Accent6,
        TableStyleType::LightStyle3,
        TableStyleType::LightStyle3Accent1,
        TableStyleType::LightStyle3Accent2,
        TableStyleType::LightStyle3Accent3,
        TableStyleType::LightStyle3Accent4,
        TableStyleType::LightStyle3Accent5,
        TableStyleType::LightStyle3Accent6,
        TableStyleType::MediumStyle1,
        TableStyleType::MediumStyle1Accent1,
        TableStyleType::MediumStyle1Accent2,
        TableStyleType::MediumStyle1Accent3,
        TableStyleType::MediumStyle1Accent4,
        TableStyleType::MediumStyle1Accent5,
        TableStyleType::MediumStyle1Accent6,
        TableStyleType::MediumStyle2,
        TableStyleType::MediumStyle2Accent1,
        TableStyleType::MediumStyle2Accent2,
        TableStyleType::MediumStyle2Accent3,
        TableStyleType::MediumStyle2Accent4,
        TableStyleType::MediumStyle2Accent5,
        TableStyleType::MediumStyle2Accent6,
        TableStyleType::MediumStyle3,
        TableStyleType::MediumStyle3Accent1,
        TableStyleType::MediumStyle3Accent2,
        TableStyleType::MediumStyle3Accent3,
        TableStyleType::MediumStyle3Accent4,
        TableStyleType::MediumStyle3Accent5,
        TableStyleType::MediumStyle3Accent6,
        TableStyleType::MediumStyle4,
        TableStyleType::MediumStyle4Accent1,
        TableStyleType::MediumStyle4Accent2,
        TableStyleType::MediumStyle4Accent3,
        TableStyleType::MediumStyle4Accent4,
        TableStyleType::MediumStyle4Accent5,
        TableStyleType::MediumStyle4Accent6,
        TableStyleType::DarkStyle1,
        TableStyleType::DarkStyle1Accent1,
        TableStyleType::DarkStyle1Accent2,
        TableStyleType::DarkStyle1Accent3,
        TableStyleType::DarkStyle1Accent4,
        TableStyleType::DarkStyle1Accent5,
        TableStyleType::DarkStyle1Accent6,
        TableStyleType::DarkStyle2,
        TableStyleType::DarkStyle2Accent1Accent2,
        TableStyleType::DarkStyle2Accent3Accent4,
        TableStyleType::DarkStyle2Accent5Accent6,
    ];

    for style_type in all_types {
        let style = get_builtin_style(style_type);
        if style.guid.eq_ignore_ascii_case(guid) {
            return Some(style);
        }
    }

    None
}

/// Parses table style options from `<a:tblPr>` XML element.
///
/// This function extracts the boolean attributes `firstRow`, `lastRow`, `firstCol`,
/// `lastCol`, `bandRow`, and `bandCol` from the table properties element.
///
/// Returns `TableStyleOptions` with the parsed values, defaulting to `false` if
/// an attribute is not present.
pub fn parse_table_style_options(xml: &str) -> TableStyleOptions {
    if xml.is_empty() {
        return TableStyleOptions::default();
    }

    let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();

    let mut options = TableStyleOptions::default();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_bytes = e.name();
                let local = local_name(name_bytes.as_ref());

                if local == "tblPr" {
                    for attr in e.attributes().flatten() {
                        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                        let value = attr.unescape_value().unwrap_or_default();

                        match key {
                            "firstRow" => options.first_row = parse_bool(&value),
                            "lastRow" => options.last_row = parse_bool(&value),
                            "firstCol" => options.first_col = parse_bool(&value),
                            "lastCol" => options.last_col = parse_bool(&value),
                            "bandRow" => options.banded_rows = parse_bool(&value),
                            "bandCol" => options.banded_cols = parse_bool(&value),
                            _ => {
                                options
                                    .unknown_attrs
                                    .push((key.to_string(), value.to_string()));
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    options
}

/// Serializes table style options to XML attributes for `<a:tblPr>`.
///
/// Returns a string containing the XML attributes that should be added to the
/// `<a:tblPr>` element. Only includes attributes that are set to `true`.
pub fn write_table_style_options(options: &TableStyleOptions) -> String {
    let mut attrs = String::new();

    if options.first_row {
        attrs.push_str(" firstRow=\"1\"");
    }
    if options.last_row {
        attrs.push_str(" lastRow=\"1\"");
    }
    if options.first_col {
        attrs.push_str(" firstCol=\"1\"");
    }
    if options.last_col {
        attrs.push_str(" lastCol=\"1\"");
    }
    if options.banded_rows {
        attrs.push_str(" bandRow=\"1\"");
    }
    if options.banded_cols {
        attrs.push_str(" bandCol=\"1\"");
    }

    // Replay unknown attributes for roundtrip fidelity.
    for (key, value) in &options.unknown_attrs {
        attrs.push(' ');
        attrs.push_str(key);
        attrs.push_str("=\"");
        attrs.push_str(value);
        attrs.push('"');
    }

    attrs
}

/// Parses a table style GUID from `<a:tableStyleId>` XML element.
///
/// Returns the GUID string if found, or `None` if the element is not present
/// or the XML is malformed.
pub fn parse_table_style_id(xml: &str) -> Option<String> {
    if xml.is_empty() {
        return None;
    }

    let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name_bytes = e.name();
                let local = local_name(name_bytes.as_ref());

                if local == "tableStyleId" {
                    // The GUID is in the text content of this element
                    buffer.clear();
                    if let Ok(Event::Text(t)) = reader.read_event_into(&mut buffer) {
                        let guid = t.xml_content().ok()?.into_owned();
                        if !guid.is_empty() {
                            return Some(guid);
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    None
}

/// Serializes a table style GUID to `<a:tableStyleId>` XML element.
///
/// Returns the complete XML element with the GUID as text content.
pub fn write_table_style_id(guid: &str) -> String {
    format!("<a:tableStyleId>{}</a:tableStyleId>", guid)
}

// ── Helper functions ──

/// Extracts the local name from a qualified XML element name.
///
/// Handles namespaced elements (e.g., "a:tblPr" -> "tblPr").
fn local_name(bytes: &[u8]) -> &str {
    let full = std::str::from_utf8(bytes).unwrap_or("");
    full.split(':').next_back().unwrap_or(full)
}

/// Parses an XML boolean value ("1", "true", "0", "false").
fn parse_bool(value: &str) -> bool {
    matches!(value, "1" | "true")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_get_builtin_style() {
        let style = get_builtin_style(TableStyleType::MediumStyle2Accent3);
        assert_eq!(style.name(), "Medium Style 2 - Accent 3");
        assert_eq!(style.guid(), "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}");
    }

    #[test]
    fn test_get_style_by_guid() {
        let style = get_style_by_guid("{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}");
        assert!(style.is_some());
        let style = style.unwrap();
        assert_eq!(style.name(), "Medium Style 2 - Accent 3");
    }

    #[test]
    fn test_get_style_by_guid_case_insensitive() {
        let style = get_style_by_guid("{f5ab1c69-6edb-4ff4-983f-18bd219ef322}");
        assert!(style.is_some());
    }

    #[test]
    fn test_get_style_by_guid_not_found() {
        let style = get_style_by_guid("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}");
        assert!(style.is_none());
    }

    #[test]
    fn test_table_style_options_default() {
        let options = TableStyleOptions::default();
        assert!(!options.first_row);
        assert!(!options.last_row);
        assert!(!options.first_col);
        assert!(!options.last_col);
        assert!(!options.banded_rows);
        assert!(!options.banded_cols);
    }

    #[test]
    fn test_table_style_options_default_pptx() {
        let options = TableStyleOptions::default_pptx();
        assert!(options.first_row);
        assert!(!options.last_row);
        assert!(!options.first_col);
        assert!(!options.last_col);
        assert!(!options.banded_rows);
        assert!(!options.banded_cols);
    }

    #[test]
    fn test_parse_table_style_options() {
        let xml = r#"<a:tblPr firstRow="1" lastRow="0" bandRow="1"/>"#;
        let options = parse_table_style_options(xml);
        assert!(options.first_row);
        assert!(!options.last_row);
        assert!(!options.first_col);
        assert!(!options.last_col);
        assert!(options.banded_rows);
        assert!(!options.banded_cols);
    }

    #[test]
    fn test_write_table_style_options() {
        let options = TableStyleOptions {
            first_row: true,
            banded_rows: true,
            ..Default::default()
        };

        let attrs = write_table_style_options(&options);
        assert!(attrs.contains("firstRow=\"1\""));
        assert!(attrs.contains("bandRow=\"1\""));
        assert!(!attrs.contains("lastRow"));
    }

    #[test]
    fn test_parse_table_style_id() {
        let xml = r#"<a:tableStyleId>{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}</a:tableStyleId>"#;
        let guid = parse_table_style_id(xml);
        assert_eq!(
            guid.as_deref(),
            Some("{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}")
        );
    }

    #[test]
    fn test_write_table_style_id() {
        let xml = write_table_style_id("{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}");
        assert_eq!(
            xml,
            "<a:tableStyleId>{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}</a:tableStyleId>"
        );
    }
}
