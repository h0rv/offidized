---
name: offidized-cli
description: Operate the Offidized CLI (`ofx`) for inspecting, creating, linting, and editing `.xlsx`, `.docx`, and `.pptx` files. Use when asked to run `ofx`, choose the right subcommand, inspect package parts, derive/apply IR, perform unified cross-format edits, or decide when to switch from one-off CLI work to Python scripting with the `offidized` package.
---

# Offidized CLI

Execute Office workflows with `ofx` first. Switch to Python only when the task becomes scripting-heavy.

## Start

- Prefer `cargo run -p offidized-cli -- <subcommand>` inside this repo so the command surface matches the checked-out source.
- Install the binary with `cargo install --path crates/offidized-cli` only when the user explicitly wants a reusable global `ofx`.
- Default to writing outputs with `-o <path>` instead of mutating inputs with `-i`.

## Choose Workflow

- Use `info`, `read`, or `part --list` to inspect a file before editing.
- Use `set`, `replace`, or `patch` for straightforward point edits.
- Use `nodes`, `capabilities`, then `edit --lint --strict` for unified edits across `xlsx`, `docx`, and `pptx`.
- Use `derive` then `apply` when the user wants a text IR workflow or wants to review/edit an intermediate representation directly.
- Use spreadsheet-specific commands (`create`, `copy-range`, `move-range`, `pivots`, `charts`, `add-chart`, `eval`, `lint`) only for `xlsx`.

## Unified Edit Guardrails

1. Run `nodes` first and treat returned IDs as canonical. Do not invent IDs.
2. Run `capabilities` when editable surfaces or style payload support is unclear.
3. Prefer `--edits-json` for more than one edit or when using `group` or typed `payload` fields.
4. Keep `--lint --strict` on by default. Use `--force` only to bypass checksum validation after confirming that is necessary.
5. If a batch spans multiple source files, use `--in-place` and do not combine it with `-o`.

## IR Guardrails

- Use `derive --mode content` by default. Use `style` or `full` only when style metadata matters.
- Use `apply --source <file>` when the IR header points at a stale or moved source path.
- Treat `apply --dry-run` as a lightweight sanity check only; it does not produce a full diff.

## Switch to Python

- Switch to `crates/offidized-py` when the task needs loops, generated documents, repeated transformations, or higher-level workbook/document/presentation construction.
- Read `references/python-bindings.md` before writing Python.
- Keep the CLI as the first choice for one-off inspection and deterministic file operations.

## References

- `references/cli-workflows.md`
- `references/python-bindings.md`
