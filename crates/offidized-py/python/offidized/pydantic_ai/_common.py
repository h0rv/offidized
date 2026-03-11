"""Shared helpers for the optional PydanticAI toolsets."""

from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path

from offidized.pydantic_ai.models import (
    ApplyResultModel,
    UnifiedEditModel,
    UnifiedEditReportModel,
)

MAX_PREVIEW_ITEMS = 100


def path_str(path: str) -> str:
    """Normalize a filesystem path to a string."""
    return str(Path(path))


def normalize_limit(limit: int, *, default: int = 20) -> int:
    """Clamp preview/list limits to a small, predictable range."""
    if limit <= 0:
        return default
    return min(limit, MAX_PREVIEW_ITEMS)


def normalize_ir_edits(edits: Iterable[UnifiedEditModel]) -> list[dict[str, str]]:
    """Coerce edit payloads into the shape expected by `UnifiedDocument`."""
    normalized: list[dict[str, str]] = []
    for edit in edits:
        item = {"id": edit.id, "text": edit.text}
        if edit.group is not None:
            item["group"] = edit.group
        normalized.append(item)
    return normalized


def apply_result_model(raw_result: object) -> ApplyResultModel:
    """Convert a native apply/save result to a Pydantic model."""
    return ApplyResultModel.model_validate(raw_result)


def edit_report_model(raw_report: object) -> UnifiedEditReportModel:
    """Convert a native edit report to a Pydantic model."""
    return UnifiedEditReportModel.model_validate(raw_report)
