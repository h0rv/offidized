# ruff: noqa: E402

"""Optional PydanticAI toolsets for offidized."""

from __future__ import annotations

from typing import Any

from offidized.pydantic_ai._compat import combine_toolsets, require_pydantic_ai
from offidized.pydantic_ai.docx import docx_toolset
from offidized.pydantic_ai.pptx import pptx_toolset
from offidized.pydantic_ai.xlsx import xlsx_toolset

require_pydantic_ai()


def compose_toolsets(*toolsets: Any) -> Any:
    """Combine multiple toolsets into one."""
    return combine_toolsets(toolsets)


def all_toolsets() -> Any:
    """Return a combined toolset for xlsx, docx, and pptx workflows."""
    return compose_toolsets(
        xlsx_toolset(),
        docx_toolset(),
        pptx_toolset(),
    )


__all__ = [
    "all_toolsets",
    "compose_toolsets",
    "docx_toolset",
    "pptx_toolset",
    "xlsx_toolset",
]
