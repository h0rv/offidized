"""Focused tests for the optional PydanticAI integration."""

from __future__ import annotations

import importlib
import sys

import pytest


def test_base_import_does_not_eagerly_load_optional_integration():
    sys.modules.pop("offidized.pydantic_ai", None)

    import offidized  # noqa: F401

    assert "offidized.pydantic_ai" not in sys.modules


def _load_optional_module():
    pytest.importorskip("pydantic_ai")
    return importlib.import_module("offidized.pydantic_ai")


def test_format_toolsets_construct_when_optional_dependency_is_installed():
    optional_module = _load_optional_module()

    xlsx = optional_module.xlsx_toolset()
    docx = optional_module.docx_toolset()
    pptx = optional_module.pptx_toolset()

    assert xlsx is not None
    assert docx is not None
    assert pptx is not None
    assert xlsx.__class__.__name__ == "FunctionToolset"
    assert docx.__class__.__name__ == "FunctionToolset"
    assert pptx.__class__.__name__ == "FunctionToolset"


def test_all_toolsets_constructs_a_combined_toolset():
    optional_module = _load_optional_module()

    combined = optional_module.all_toolsets()

    assert combined is not None
    assert combined.__class__.__name__ == "CombinedToolset"


def test_compose_toolsets_constructs_a_combined_toolset():
    optional_module = _load_optional_module()

    combined = optional_module.compose_toolsets(
        optional_module.xlsx_toolset(),
        optional_module.docx_toolset(),
    )

    assert combined is not None
    assert combined.__class__.__name__ == "CombinedToolset"
