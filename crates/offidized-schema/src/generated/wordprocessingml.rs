#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct WordprocessingmlElementDescriptor {
    pub schema_path: &'static str,
    pub class_name: &'static str,
    pub qualified_name: &'static str,
}

pub const ELEMENTS: &[WordprocessingmlElementDescriptor] = &[
    WordprocessingmlElementDescriptor {
        schema_path: "/w:document",
        class_name: "Document",
        qualified_name: "w:document",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:document/w:body",
        class_name: "Body",
        qualified_name: "w:body",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:document/w:body/w:p",
        class_name: "Paragraph",
        qualified_name: "w:p",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:document/w:body/w:p/w:r",
        class_name: "Run",
        qualified_name: "w:r",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:document/w:body/w:p/w:r/w:t",
        class_name: "Text",
        qualified_name: "w:t",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:document/w:body/w:tbl",
        class_name: "Table",
        qualified_name: "w:tbl",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:document/w:body/w:tbl/w:tr",
        class_name: "TableRow",
        qualified_name: "w:tr",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:document/w:body/w:tbl/w:tr/w:tc",
        class_name: "TableCell",
        qualified_name: "w:tc",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:styles",
        class_name: "Styles",
        qualified_name: "w:styles",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:styles/w:style",
        class_name: "Style",
        qualified_name: "w:style",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:numbering",
        class_name: "Numbering",
        qualified_name: "w:numbering",
    },
    WordprocessingmlElementDescriptor {
        schema_path: "/w:numbering/w:num",
        class_name: "NumberingInstance",
        qualified_name: "w:num",
    },
];
