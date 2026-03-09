---
name: offidized-mcp
description: Operate the Offidized MCP server for deterministic editing of xlsx, docx, and pptx files. Use when asked to upload Office files, discover stable target IDs, lint edit batches, run dry-run previews, apply edits safely, or troubleshoot MCP session/auth/deploy issues.
---

# Offidized MCP

Execute Office edits through the MCP workflow.

## Execute Workflow

1. Call `workspace_put` to upload bytes.
2. Call `workspace_targets` to discover canonical IDs.
3. Call `workspace_lint_edits` to catch invalid targets/anchors.
4. Call `workspace_preview_edits` with `strict=true` for dry-run validation.
5. Call `workspace_apply_edits` only after preview is clean.
6. Call `resources/read` with returned `resource_uri` to retrieve output bytes.

Treat `workspace_targets` IDs as source of truth. Do not invent IDs.

## Build Edit Payloads

Use `edits` entries with:

- `id` (required)
- `text` (optional)
- `group` (optional)
- `payload` (optional typed object)

Use typed payloads for style metadata:

- `xlsx_cell_style`: `bold`, `italic`, `number_format`
- `pptx_text_style`: `bold`, `italic`, `font_size`, `font_color`, `font_name`

## Enforce Safety

Default to `lint=true` and `strict=true`.
If strict blocks apply, inspect diagnostics and regenerate edits from `workspace_targets`.

## References

- `references/workflow-examples.md`
- `references/troubleshooting.md`
