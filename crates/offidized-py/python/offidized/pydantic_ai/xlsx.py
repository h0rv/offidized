"""PydanticAI toolset helpers for `.xlsx` files."""

from __future__ import annotations

from offidized import UnifiedDocument, Workbook, ir_apply, ir_derive
from offidized.pydantic_ai._common import (
    apply_result_model,
    edit_report_model,
    normalize_ir_edits,
    path_str,
)
from offidized.pydantic_ai._compat import new_function_toolset
from offidized.pydantic_ai.models import (
    ApplyIrEditsResponse,
    ApplyIrResponse,
    IrTextResponse,
    XlsxApplyIrEditsInput,
    XlsxApplyIrInput,
    XlsxDeriveIrInput,
    XlsxListSheetsInput,
    XlsxListSheetsResponse,
    XlsxOpenSummaryInput,
    XlsxOpenSummaryResponse,
    XlsxReadCellInput,
    XlsxReadCellResponse,
    XlsxSetFormulaInput,
    XlsxSetFormulaResponse,
    XlsxWriteCellInput,
    XlsxWriteCellResponse,
)


def xlsx_open_summary(request: XlsxOpenSummaryInput) -> XlsxOpenSummaryResponse:
    """Return the workbook path and worksheet names."""
    workbook = Workbook.open(request.path)
    sheet_names = workbook.sheet_names()
    return XlsxOpenSummaryResponse(
        path=path_str(request.path),
        sheet_count=len(sheet_names),
        sheet_names=sheet_names,
    )


def xlsx_list_sheets(request: XlsxListSheetsInput) -> XlsxListSheetsResponse:
    """List worksheet names in a workbook."""
    workbook = Workbook.open(request.path)
    return XlsxListSheetsResponse(sheet_names=workbook.sheet_names())


def xlsx_read_cell(request: XlsxReadCellInput) -> XlsxReadCellResponse:
    """Read a cell value from a worksheet."""
    workbook = Workbook.open(request.path)
    worksheet = workbook.sheet(request.sheet_name)
    return XlsxReadCellResponse(value=worksheet.cell_value(request.cell_reference))


def xlsx_write_cell(request: XlsxWriteCellInput) -> XlsxWriteCellResponse:
    """Write a single cell value and save the workbook."""
    workbook = Workbook.open(request.source_path)
    worksheet = workbook.sheet(request.sheet_name)
    worksheet.set_cell_value(request.cell_reference, request.value)
    workbook.save(request.output_path)
    return XlsxWriteCellResponse(
        source_path=path_str(request.source_path),
        output_path=path_str(request.output_path),
        sheet_name=request.sheet_name,
        cell_reference=request.cell_reference,
        value=request.value,
    )


def xlsx_set_formula(request: XlsxSetFormulaInput) -> XlsxSetFormulaResponse:
    """Write a single formula and save the workbook."""
    workbook = Workbook.open(request.source_path)
    worksheet = workbook.sheet(request.sheet_name)
    worksheet.set_cell_formula(request.cell_reference, request.formula)
    workbook.save(request.output_path)
    return XlsxSetFormulaResponse(
        source_path=path_str(request.source_path),
        output_path=path_str(request.output_path),
        sheet_name=request.sheet_name,
        cell_reference=request.cell_reference,
        formula=request.formula,
    )


def xlsx_derive_ir(request: XlsxDeriveIrInput) -> IrTextResponse:
    """Derive IR text for an `.xlsx` file."""
    return IrTextResponse(
        ir_text=ir_derive(
            request.path,
            mode=request.mode,
            sheet=request.sheet_name,
            range=request.range_reference,
        )
    )


def xlsx_apply_ir(request: XlsxApplyIrInput) -> ApplyIrResponse:
    """Apply IR text to an `.xlsx` file and save the result."""
    apply_result = ir_apply(
        request.ir_text,
        request.output_path,
        source_override=request.source_path,
        force=request.force,
    )
    return ApplyIrResponse(
        source_path=path_str(request.source_path),
        output_path=path_str(request.output_path),
        apply_result=apply_result_model(apply_result),
    )


def xlsx_apply_ir_edits(request: XlsxApplyIrEditsInput) -> ApplyIrEditsResponse:
    """Apply structured IR edits to an `.xlsx` file."""
    document = UnifiedDocument.derive(
        request.source_path,
        mode=request.mode,
        sheet=request.sheet_name,
        range=request.range_reference,
    )
    edit_report = document.apply_edits(normalize_ir_edits(request.edits))
    save_result = document.save_as(
        request.output_path,
        source_override=request.source_path,
        force=request.force,
    )
    return ApplyIrEditsResponse(
        source_path=path_str(request.source_path),
        output_path=path_str(request.output_path),
        edit_report=edit_report_model(edit_report),
        save_result=apply_result_model(save_result),
    )


def xlsx_toolset() -> object:
    """Create a PydanticAI toolset for `.xlsx` workflows."""
    return new_function_toolset(
        [
            xlsx_open_summary,
            xlsx_list_sheets,
            xlsx_read_cell,
            xlsx_write_cell,
            xlsx_set_formula,
            xlsx_derive_ir,
            xlsx_apply_ir,
            xlsx_apply_ir_edits,
        ]
    )
