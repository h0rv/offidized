"""offidized — Rust-native OOXML library for xlsx, docx, and pptx."""

from offidized._native import (
    # Errors
    OffidizedError,
    OffidizedIoError,
    OffidizedRuntimeError,
    OffidizedUnsupportedError,
    OffidizedValueError,
    # xlsx — core
    Workbook,
    Worksheet,
    XlsxCell,
    XlsxStyle,
    XlsxFont,
    XlsxFill,
    XlsxBorder,
    XlsxAlignment,
    # xlsx — row/column
    XlsxRow,
    XlsxColumn,
    # xlsx — comment
    XlsxComment,
    # xlsx — protection
    XlsxSheetProtection,
    XlsxWorkbookProtection,
    # xlsx — page setup
    XlsxPageSetup,
    XlsxPageMargins,
    XlsxPrintHeaderFooter,
    XlsxPageBreaks,
    XlsxSheetViewOptions,
    # xlsx — image
    XlsxWorksheetImage,
    # xlsx — data validation
    XlsxDataValidation,
    # xlsx — table
    XlsxTableColumn,
    XlsxWorksheetTable,
    # xlsx — sparkline
    XlsxSparkline,
    XlsxSparklineGroup,
    # xlsx — conditional formatting
    XlsxConditionalFormatting,
    # xlsx — chart
    XlsxChartDataRef,
    XlsxChartSeries,
    XlsxChartAxis,
    XlsxChartLegend,
    XlsxChart,
    # xlsx — pivot table
    XlsxPivotField,
    XlsxPivotDataField,
    XlsxPivotTable,
    # xlsx — lint
    XlsxLintFinding,
    XlsxLintReport,
    # xlsx — rich text
    XlsxRichTextRun,
    # docx
    Document,
    DocxParagraph,
    DocxRun,
    DocxTable,
    DocxTableCell,
    DocxSection,
    DocxDocumentProperties,
    # pptx
    Presentation,
    PresentationProperties,
    Slide,
    Table,
    Chart,
    Image,
    Shape,
    ShapeParagraph,
    TextRun,
    SlideShowSettings,
    CustomShow,
    SlideTransition,
    # ir
    UnifiedDocument,
    ir_derive,
    ir_apply,
    ir_derive_from_bytes,
    ir_apply_to_bytes,
)

__all__ = [
    # Errors
    "OffidizedError",
    "OffidizedIoError",
    "OffidizedRuntimeError",
    "OffidizedUnsupportedError",
    "OffidizedValueError",
    # xlsx — core
    "Workbook",
    "Worksheet",
    "XlsxCell",
    "XlsxStyle",
    "XlsxFont",
    "XlsxFill",
    "XlsxBorder",
    "XlsxAlignment",
    # xlsx — row/column
    "XlsxRow",
    "XlsxColumn",
    # xlsx — comment
    "XlsxComment",
    # xlsx — protection
    "XlsxSheetProtection",
    "XlsxWorkbookProtection",
    # xlsx — page setup
    "XlsxPageSetup",
    "XlsxPageMargins",
    "XlsxPrintHeaderFooter",
    "XlsxPageBreaks",
    "XlsxSheetViewOptions",
    # xlsx — image
    "XlsxWorksheetImage",
    # xlsx — data validation
    "XlsxDataValidation",
    # xlsx — table
    "XlsxTableColumn",
    "XlsxWorksheetTable",
    # xlsx — sparkline
    "XlsxSparkline",
    "XlsxSparklineGroup",
    # xlsx — conditional formatting
    "XlsxConditionalFormatting",
    # xlsx — chart
    "XlsxChartDataRef",
    "XlsxChartSeries",
    "XlsxChartAxis",
    "XlsxChartLegend",
    "XlsxChart",
    # xlsx — pivot table
    "XlsxPivotField",
    "XlsxPivotDataField",
    "XlsxPivotTable",
    # xlsx — lint
    "XlsxLintFinding",
    "XlsxLintReport",
    # xlsx — rich text
    "XlsxRichTextRun",
    # docx
    "Document",
    "DocxParagraph",
    "DocxRun",
    "DocxTable",
    "DocxTableCell",
    "DocxSection",
    "DocxDocumentProperties",
    # pptx
    "Presentation",
    "PresentationProperties",
    "Slide",
    "Table",
    "Chart",
    "Image",
    "Shape",
    "ShapeParagraph",
    "TextRun",
    "SlideShowSettings",
    "CustomShow",
    "SlideTransition",
    # ir
    "UnifiedDocument",
    "ir_derive",
    "ir_apply",
    "ir_derive_from_bytes",
    "ir_apply_to_bytes",
]
