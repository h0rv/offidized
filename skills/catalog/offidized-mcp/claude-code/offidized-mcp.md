# offidized-mcp

Use the Offidized MCP workflow for deterministic Office edits.

## Procedure

1. Upload file bytes.
2. Discover canonical IDs via `workspace_targets`.
3. Validate edits via `workspace_lint_edits`.
4. Dry-run via `workspace_preview_edits` (`strict=true`).
5. Apply via `workspace_apply_edits` only if clean.
6. Read output bytes with `resources/read` on returned `resource_uri`.
7. Download updated bytes.

## Guardrails

- Never invent IDs; only use IDs from `workspace_targets`.
- Keep `lint=true` and `strict=true` by default.
- If strict fails, fix diagnostics first.
