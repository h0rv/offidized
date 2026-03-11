"""Compatibility helpers for the optional PydanticAI integration."""

from __future__ import annotations

from typing import Any, Iterable

MISSING_DEPENDENCY_MESSAGE = (
    "offidized.pydantic_ai requires the optional 'pydantic-ai' dependency. "
    "Install it with `pip install offidized[pydantic-ai]`. "
    "The current PydanticAI release line requires Python 3.10+."
)


def require_pydantic_ai() -> tuple[Any, Any]:
    """Import and return the PydanticAI toolset classes."""
    try:
        from pydantic_ai import CombinedToolset, FunctionToolset
    except ImportError as exc:  # pragma: no cover - exercised in import tests
        raise ImportError(MISSING_DEPENDENCY_MESSAGE) from exc
    return FunctionToolset, CombinedToolset


def new_function_toolset(tools: Iterable[Any]) -> Any:
    """Create a function toolset from plain Python callables."""
    function_toolset, _ = require_pydantic_ai()
    return function_toolset(list(tools))


def combine_toolsets(toolsets: Iterable[Any]) -> Any:
    """Combine multiple toolsets into a single toolset."""
    _, combined_toolset = require_pydantic_ai()
    return combined_toolset(list(toolsets))
