//! # offidized-xlsx
//!
//! High-level Excel API.
//!
//! ```ignore
//! use offidized_xlsx::Workbook;
//!
//! let mut wb = Workbook::new();
//! let ws = wb.add_sheet("Sales");
//! ws.cell_mut("A1")?.set_value("Product");
//! ws.cell_mut("B1")?.set_value("Revenue");
//! ws.cell_mut("A2")?.set_value("Widget");
//! ws.cell_mut("B2")?.set_value(42_000);
//! ws.cell_mut("B2")?.set_style_id(1);
//! wb.save("output.xlsx")?;
//!
//! let mut wb = Workbook::open("output.xlsx")?;
//! if let Some(ws) = wb.sheet_mut("Sales") {
//!     ws.cell_mut("B3")?.set_value(58_000);
//!     ws.cell_mut("B4")?.set_formula("SUM(B2:B3)");
//! }
//! wb.save("output_updated.xlsx")?;
//! # Ok::<(), offidized_xlsx::XlsxError>(())
//! ```

pub mod auto_filter;
pub mod cell;
pub mod chart;
pub(crate) mod chart_io;
pub mod color;
pub mod column;
pub mod error;
pub mod finance;
pub(crate) mod formula_bridge;
pub mod named_style;
pub mod numfmt;
pub mod pivot_table;
pub(crate) mod pivot_table_io;
pub mod print_settings;
pub mod range;
pub mod reference;
pub mod row;
pub mod shared_strings;
pub mod sparkline;
pub mod style;
pub mod theme;
pub mod workbook;
pub mod workbook_lint;
pub mod worksheet;

pub use auto_filter::{AutoFilter, CustomFilter, CustomFilterOperator, FilterColumn, FilterType};
pub use cell::{
    is_builtin_date_format, Cell, CellComment, CellValue, RichTextRun, BUILTIN_DATE_FORMAT_IDS,
};
pub use chart::{
    BarDirection, Chart, ChartAxis, ChartDataRef, ChartGrouping, ChartLegend, ChartSeries,
    ChartType,
};
pub use color::{apply_tint, resolve_color, DEFAULT_THEME_COLORS, INDEXED_COLORS};
pub use error::{Result, XlsxError};
pub use finance::{
    ChartTemplate, FinFormat, FinanceChartTemplateBuilder, FinanceModelBuilder, MeasureType,
    WorkbookPivotBuilder,
};
pub use named_style::{builtin_style_name, CellStyleXf, NamedStyle};
pub use numfmt::{
    compile_format, format_value, format_value_compiled, get_builtin_format, is_date_format,
    CompiledFormat, FormattedValue,
};
pub use pivot_table::{
    PivotDataField, PivotField, PivotFieldSort, PivotSourceReference, PivotSubtotalFunction,
    PivotTable,
};
pub use print_settings::{PageBreak, PageBreaks, PrintArea, PrintHeaderFooter};
pub use reference::{a1_to_r1c1, r1c1_to_a1};
pub use shared_strings::SharedStringEntry;
pub use sparkline::{
    Sparkline, SparklineAxisType, SparklineColors, SparklineEmptyCells, SparklineGroup,
    SparklineType,
};
pub use style::{
    Alignment, Border, BorderSide, CellProtection, ColorReference, Fill, Font, FontScheme,
    FontVerticalAlign, GradientFill, GradientFillType, GradientStop, HorizontalAlignment,
    PatternFill, PatternFillType, Style, StyleTable, ThemeColor, VerticalAlignment,
};
pub use theme::{ParsedTheme, DEFAULT_THEME_COLORS as THEME_DEFAULTS};
pub use workbook::{CalculationSettings, DefinedName, Workbook, WorkbookProtection};
pub use workbook_lint::{LintFinding, LintLocation, LintReport, LintSeverity, WorkbookLintBuilder};
pub use worksheet::{
    avg, sum, CellAnchor, CfValueObject, CfValueObjectType, ColorScaleStop, Comment, CommentReply,
    ConditionalFormatting, ConditionalFormattingOperator, ConditionalFormattingRuleType,
    DataValidation, DataValidationErrorStyle, DataValidationType, FreezePane, Hyperlink,
    ImageAnchorType, PageMargins, PageOrientation, PageSetup, PivotBuilder, PivotValueSpec,
    SheetProtection, SheetViewOptions, SheetVisibility, TableColumn, TotalFunction, Worksheet,
    WorksheetImage, WorksheetImageExt, WorksheetTable,
};
