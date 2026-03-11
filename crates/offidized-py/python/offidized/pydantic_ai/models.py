"""Pydantic models for the optional offidized PydanticAI toolsets."""

from __future__ import annotations

from typing import Any, Optional

from offidized.pydantic_ai._compat import MISSING_DEPENDENCY_MESSAGE

try:
    from pydantic import BaseModel
except ImportError as exc:  # pragma: no cover - exercised in import tests
    raise ImportError(MISSING_DEPENDENCY_MESSAGE) from exc


class ApplyResultModel(BaseModel):
    """Structured result returned by IR apply/save operations."""

    cells_updated: int
    cells_created: int
    cells_cleared: int
    charts_added: int
    warnings: list[str]


class UnifiedDiagnosticModel(BaseModel):
    """Structured diagnostic returned by IR lint/edit operations."""

    severity: str
    code: str
    message: str
    id: Optional[str] = None


class UnifiedEditModel(BaseModel):
    """Single structured IR edit."""

    id: str
    text: str
    group: Optional[str] = None


class UnifiedEditReportModel(BaseModel):
    """Structured edit report returned by `UnifiedDocument.apply_edits`."""

    requested: int
    applied: int
    skipped: int
    diagnostics: list[UnifiedDiagnosticModel]


class IrTextResponse(BaseModel):
    """Structured IR text response."""

    ir_text: str


class ApplyIrResponse(BaseModel):
    """Structured IR apply response."""

    source_path: str
    output_path: str
    apply_result: ApplyResultModel


class ApplyIrEditsResponse(BaseModel):
    """Structured IR edit-apply response."""

    source_path: str
    output_path: str
    edit_report: UnifiedEditReportModel
    save_result: ApplyResultModel


class XlsxOpenSummaryInput(BaseModel):
    path: str


class XlsxOpenSummaryResponse(BaseModel):
    path: str
    sheet_count: int
    sheet_names: list[str]


class XlsxListSheetsInput(BaseModel):
    path: str


class XlsxListSheetsResponse(BaseModel):
    sheet_names: list[str]


class XlsxReadCellInput(BaseModel):
    path: str
    sheet_name: str
    cell_reference: str


class XlsxReadCellResponse(BaseModel):
    value: Any


class XlsxWriteCellInput(BaseModel):
    source_path: str
    output_path: str
    sheet_name: str
    cell_reference: str
    value: Any


class XlsxWriteCellResponse(BaseModel):
    source_path: str
    output_path: str
    sheet_name: str
    cell_reference: str
    value: Any


class XlsxSetFormulaInput(BaseModel):
    source_path: str
    output_path: str
    sheet_name: str
    cell_reference: str
    formula: str


class XlsxSetFormulaResponse(BaseModel):
    source_path: str
    output_path: str
    sheet_name: str
    cell_reference: str
    formula: str


class XlsxDeriveIrInput(BaseModel):
    path: str
    mode: str = "content"
    sheet_name: Optional[str] = None
    range_reference: Optional[str] = None


class XlsxApplyIrInput(BaseModel):
    source_path: str
    output_path: str
    ir_text: str
    force: bool = False


class XlsxApplyIrEditsInput(BaseModel):
    source_path: str
    output_path: str
    edits: list[UnifiedEditModel]
    mode: str = "content"
    sheet_name: Optional[str] = None
    range_reference: Optional[str] = None
    force: bool = False


class DocxOpenSummaryInput(BaseModel):
    path: str
    paragraph_limit: int = 10


class DocxOpenSummaryResponse(BaseModel):
    path: str
    paragraph_count: int
    table_count: int
    image_count: int
    paragraph_preview: list[str]


class DocxParagraphModel(BaseModel):
    index: int
    text: str
    run_count: int


class DocxListParagraphsInput(BaseModel):
    path: str
    limit: int = 50


class DocxListParagraphsResponse(BaseModel):
    paragraphs: list[DocxParagraphModel]


class DocxAddParagraphInput(BaseModel):
    source_path: str
    output_path: str
    text: str
    style_id: Optional[str] = None


class DocxAddParagraphResponse(BaseModel):
    source_path: str
    output_path: str
    paragraph_text: str
    style_id: Optional[str] = None


class DocxSetParagraphTextInput(BaseModel):
    source_path: str
    output_path: str
    paragraph_index: int
    text: str


class DocxSetParagraphTextResponse(BaseModel):
    source_path: str
    output_path: str
    paragraph_index: int
    text: str


class DocxDeriveIrInput(BaseModel):
    path: str
    mode: str = "content"


class DocxApplyIrInput(BaseModel):
    source_path: str
    output_path: str
    ir_text: str
    force: bool = False


class DocxApplyIrEditsInput(BaseModel):
    source_path: str
    output_path: str
    edits: list[UnifiedEditModel]
    mode: str = "content"
    force: bool = False


class PptxOpenSummaryInput(BaseModel):
    path: str
    slide_limit: int = 10


class PptxSlideModel(BaseModel):
    index: int
    title: Optional[str] = None
    shape_count: int
    table_count: int
    chart_count: int
    notes: Optional[str] = None


class PptxOpenSummaryResponse(BaseModel):
    path: str
    slide_count: int
    slides: list[PptxSlideModel]


class PptxListSlidesInput(BaseModel):
    path: str
    limit: int = 50


class PptxListSlidesResponse(BaseModel):
    slides: list[PptxSlideModel]


class PptxAddSlideWithTitleInput(BaseModel):
    source_path: str
    output_path: str
    title: str


class PptxAddSlideWithTitleResponse(BaseModel):
    source_path: str
    output_path: str
    slide_index: int
    title: str


class PptxAddTextShapeInput(BaseModel):
    source_path: str
    output_path: str
    slide_index: int
    shape_name: str
    text: str


class PptxAddTextShapeResponse(BaseModel):
    source_path: str
    output_path: str
    slide_index: int
    shape_index: int
    shape_name: str
    text: str


class PptxReplaceTextInput(BaseModel):
    source_path: str
    output_path: str
    old_text: str
    new_text: str


class PptxReplaceTextResponse(BaseModel):
    source_path: str
    output_path: str
    replacements: int


class PptxDeriveIrInput(BaseModel):
    path: str
    mode: str = "content"


class PptxApplyIrInput(BaseModel):
    source_path: str
    output_path: str
    ir_text: str
    force: bool = False


class PptxApplyIrEditsInput(BaseModel):
    source_path: str
    output_path: str
    edits: list[UnifiedEditModel]
    mode: str = "content"
    force: bool = False
