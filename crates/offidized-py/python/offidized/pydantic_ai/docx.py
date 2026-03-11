"""PydanticAI toolset helpers for `.docx` files."""

from __future__ import annotations

from offidized import Document, UnifiedDocument, ir_apply, ir_derive
from offidized.pydantic_ai._common import (
    apply_result_model,
    edit_report_model,
    normalize_ir_edits,
    normalize_limit,
    path_str,
)
from offidized.pydantic_ai._compat import new_function_toolset
from offidized.pydantic_ai.models import (
    ApplyIrEditsResponse,
    ApplyIrResponse,
    DocxAddParagraphInput,
    DocxAddParagraphResponse,
    DocxApplyIrEditsInput,
    DocxApplyIrInput,
    DocxDeriveIrInput,
    DocxListParagraphsInput,
    DocxListParagraphsResponse,
    DocxOpenSummaryInput,
    DocxOpenSummaryResponse,
    DocxParagraphModel,
    DocxSetParagraphTextInput,
    DocxSetParagraphTextResponse,
    IrTextResponse,
)


def docx_open_summary(request: DocxOpenSummaryInput) -> DocxOpenSummaryResponse:
    """Return document counts and a short paragraph preview."""
    document = Document.open(request.path)
    limit = normalize_limit(request.paragraph_limit, default=10)
    paragraphs = document.paragraphs()
    return DocxOpenSummaryResponse(
        path=path_str(request.path),
        paragraph_count=document.paragraph_count(),
        table_count=document.table_count(),
        image_count=document.image_count(),
        paragraph_preview=[paragraph.text() for paragraph in paragraphs[:limit]],
    )


def docx_list_paragraphs(
    request: DocxListParagraphsInput,
) -> DocxListParagraphsResponse:
    """List document paragraphs with indexes and run counts."""
    document = Document.open(request.path)
    items: list[DocxParagraphModel] = []
    limit = normalize_limit(request.limit, default=50)
    for index, paragraph in enumerate(document.paragraphs()[:limit]):
        items.append(
            DocxParagraphModel(
                index=index,
                text=paragraph.text(),
                run_count=paragraph.run_count(),
            )
        )
    return DocxListParagraphsResponse(paragraphs=items)


def docx_add_paragraph(request: DocxAddParagraphInput) -> DocxAddParagraphResponse:
    """Add a paragraph and save the document."""
    document = Document.open(request.source_path)
    if request.style_id:
        paragraph = document.add_paragraph_with_style(request.text, request.style_id)
    else:
        paragraph = document.add_paragraph(request.text)
    document.save(request.output_path)
    return DocxAddParagraphResponse(
        source_path=path_str(request.source_path),
        output_path=path_str(request.output_path),
        paragraph_text=paragraph.text(),
        style_id=request.style_id,
    )


def docx_set_paragraph_text(
    request: DocxSetParagraphTextInput,
) -> DocxSetParagraphTextResponse:
    """Replace one paragraph's text and save the document."""
    document = Document.open(request.source_path)
    paragraph = document.paragraph(request.paragraph_index)
    paragraph.set_text(request.text)
    document.save(request.output_path)
    return DocxSetParagraphTextResponse(
        source_path=path_str(request.source_path),
        output_path=path_str(request.output_path),
        paragraph_index=request.paragraph_index,
        text=paragraph.text(),
    )


def docx_derive_ir(request: DocxDeriveIrInput) -> IrTextResponse:
    """Derive IR text for a `.docx` file."""
    return IrTextResponse(ir_text=ir_derive(request.path, mode=request.mode))


def docx_apply_ir(request: DocxApplyIrInput) -> ApplyIrResponse:
    """Apply IR text to a `.docx` file and save the result."""
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


def docx_apply_ir_edits(request: DocxApplyIrEditsInput) -> ApplyIrEditsResponse:
    """Apply structured IR edits to a `.docx` file."""
    document = UnifiedDocument.derive(request.source_path, mode=request.mode)
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


def docx_toolset() -> object:
    """Create a PydanticAI toolset for `.docx` workflows."""
    return new_function_toolset(
        [
            docx_open_summary,
            docx_list_paragraphs,
            docx_add_paragraph,
            docx_set_paragraph_text,
            docx_derive_ir,
            docx_apply_ir,
            docx_apply_ir_edits,
        ]
    )
