#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SpreadsheetmlElementDescriptor {
    pub schema_path: &'static str,
    pub class_name: &'static str,
    pub qualified_name: &'static str,
}

pub const ELEMENTS: &[SpreadsheetmlElementDescriptor] = &[
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:workbook",
        class_name: "Workbook",
        qualified_name: "x:workbook",
    },
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:workbook/x:sheets",
        class_name: "Sheets",
        qualified_name: "x:sheets",
    },
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:workbook/x:sheets/x:sheet",
        class_name: "Sheet",
        qualified_name: "x:sheet",
    },
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:worksheet",
        class_name: "Worksheet",
        qualified_name: "x:worksheet",
    },
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:worksheet/x:sheetData",
        class_name: "SheetData",
        qualified_name: "x:sheetData",
    },
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:worksheet/x:sheetData/x:row",
        class_name: "Row",
        qualified_name: "x:row",
    },
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:worksheet/x:sheetData/x:row/x:c",
        class_name: "Cell",
        qualified_name: "x:c",
    },
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:worksheet/x:sheetData/x:row/x:c/x:v",
        class_name: "CellValue",
        qualified_name: "x:v",
    },
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:sst",
        class_name: "SharedStringTable",
        qualified_name: "x:sst",
    },
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:sst/x:si",
        class_name: "SharedStringItem",
        qualified_name: "x:si",
    },
    SpreadsheetmlElementDescriptor {
        schema_path: "/x:styleSheet",
        class_name: "Stylesheet",
        qualified_name: "x:styleSheet",
    },
];
