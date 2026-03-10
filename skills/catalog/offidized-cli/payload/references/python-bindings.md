# Python Bindings

Use Python as the second choice after `ofx`, not the first. Switch when the task needs real scripting or higher-level object APIs.

## Install

- Published package: `pip install offidized`
- Local examples from this repo: `cd crates/offidized-py && uv run python examples/xlsx/beyond_openpyxl.py`

The import name is `offidized`.

## When To Prefer Python

- Generate workbooks, documents, or presentations from data
- Repeat the same edit pattern across many files or many targets
- Build richer layouts than a one-off CLI command can express cleanly
- Stay in Python for an existing automation script or notebook

## Where To Read First

- Package overview: `crates/offidized-py/README.md`
- Import and smoke coverage: `crates/offidized-py/tests/test_smoke.py`
- Excel examples: `crates/offidized-py/examples/xlsx/`
- Word examples: `crates/offidized-py/examples/docx/`
- PowerPoint examples: `crates/offidized-py/examples/pptx/`

## Fast Local Commands

- Run one example:

```bash
cd crates/offidized-py
uv run python examples/docx/01_styled_document.py
```

- Run smoke tests:

```bash
cd crates/offidized-py
uv run pytest tests/test_smoke.py -q
```

## Example Entry Points

- `examples/xlsx/beyond_openpyxl.py`: feature-heavy spreadsheet construction and roundtrip fidelity
- `examples/xlsx/quant_portfolio.py`: a more realistic workbook workflow
- `examples/docx/01_styled_document.py`: styled document creation
- `examples/docx/07_contract.py`: longer-form document generation
- `examples/pptx/07_pitch_deck.py`: deck generation

## Decision Rule

Stay on the CLI when a few deterministic commands can solve the task quickly.
Switch to Python when code is the clearer representation of the requested workflow.
