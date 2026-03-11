"""PydanticAI toolset helpers for `.pptx` files."""

from __future__ import annotations

from offidized import Presentation, UnifiedDocument, ir_apply, ir_derive
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
    IrTextResponse,
    PptxAddSlideWithTitleInput,
    PptxAddSlideWithTitleResponse,
    PptxAddTextShapeInput,
    PptxAddTextShapeResponse,
    PptxApplyIrEditsInput,
    PptxApplyIrInput,
    PptxDeriveIrInput,
    PptxListSlidesInput,
    PptxListSlidesResponse,
    PptxOpenSummaryInput,
    PptxOpenSummaryResponse,
    PptxReplaceTextInput,
    PptxReplaceTextResponse,
    PptxSlideModel,
)


def pptx_open_summary(request: PptxOpenSummaryInput) -> PptxOpenSummaryResponse:
    """Return presentation counts and a short slide preview."""
    presentation = Presentation.open(request.path)
    limit = normalize_limit(request.slide_limit, default=10)
    slides: list[PptxSlideModel] = []
    for index in range(min(presentation.slide_count(), limit)):
        slide = presentation.get_slide(index)
        slides.append(
            PptxSlideModel(
                index=index,
                title=slide.title(),
                shape_count=slide.shape_count(),
                table_count=slide.table_count(),
                chart_count=slide.chart_count(),
                notes=slide.notes(),
            )
        )
    return PptxOpenSummaryResponse(
        path=path_str(request.path),
        slide_count=presentation.slide_count(),
        slides=slides,
    )


def pptx_list_slides(request: PptxListSlidesInput) -> PptxListSlidesResponse:
    """List slides with titles and counts."""
    presentation = Presentation.open(request.path)
    items: list[PptxSlideModel] = []
    limit = normalize_limit(request.limit, default=50)
    for index in range(min(presentation.slide_count(), limit)):
        slide = presentation.get_slide(index)
        items.append(
            PptxSlideModel(
                index=index,
                title=slide.title(),
                shape_count=slide.shape_count(),
                table_count=slide.table_count(),
                chart_count=slide.chart_count(),
                notes=slide.notes(),
            )
        )
    return PptxListSlidesResponse(slides=items)


def pptx_add_slide_with_title(
    request: PptxAddSlideWithTitleInput,
) -> PptxAddSlideWithTitleResponse:
    """Add a titled slide and save the presentation."""
    presentation = Presentation.open(request.source_path)
    slide_index = presentation.add_slide_with_title(request.title)
    presentation.save(request.output_path)
    return PptxAddSlideWithTitleResponse(
        source_path=path_str(request.source_path),
        output_path=path_str(request.output_path),
        slide_index=slide_index,
        title=request.title,
    )


def pptx_add_text_shape(request: PptxAddTextShapeInput) -> PptxAddTextShapeResponse:
    """Add a text shape to a slide and save the presentation."""
    presentation = Presentation.open(request.source_path)
    slide = presentation.get_slide(request.slide_index)
    shape_index = slide.add_shape(request.shape_name)
    shape = slide.get_shape(shape_index)
    shape.add_paragraph_with_text(request.text)
    presentation.save(request.output_path)
    return PptxAddTextShapeResponse(
        source_path=path_str(request.source_path),
        output_path=path_str(request.output_path),
        slide_index=request.slide_index,
        shape_index=shape_index,
        shape_name=request.shape_name,
        text=request.text,
    )


def pptx_replace_text(request: PptxReplaceTextInput) -> PptxReplaceTextResponse:
    """Replace text throughout a presentation and save it."""
    presentation = Presentation.open(request.source_path)
    replacements = presentation.replace_text(request.old_text, request.new_text)
    presentation.save(request.output_path)
    return PptxReplaceTextResponse(
        source_path=path_str(request.source_path),
        output_path=path_str(request.output_path),
        replacements=replacements,
    )


def pptx_derive_ir(request: PptxDeriveIrInput) -> IrTextResponse:
    """Derive IR text for a `.pptx` file."""
    return IrTextResponse(ir_text=ir_derive(request.path, mode=request.mode))


def pptx_apply_ir(request: PptxApplyIrInput) -> ApplyIrResponse:
    """Apply IR text to a `.pptx` file and save the result."""
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


def pptx_apply_ir_edits(request: PptxApplyIrEditsInput) -> ApplyIrEditsResponse:
    """Apply structured IR edits to a `.pptx` file."""
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


def pptx_toolset() -> object:
    """Create a PydanticAI toolset for `.pptx` workflows."""
    return new_function_toolset(
        [
            pptx_open_summary,
            pptx_list_slides,
            pptx_add_slide_with_title,
            pptx_add_text_shape,
            pptx_replace_text,
            pptx_derive_ir,
            pptx_apply_ir,
            pptx_apply_ir_edits,
        ]
    )
