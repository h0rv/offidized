# Workflow Examples

## Discover targets

Tool: `workspace_targets`

```json
{
  "filename": "earnings.xlsx",
  "mode": "content",
  "limit": 200
}
```

## Lint edits

Tool: `workspace_lint_edits`

```json
{
  "filename": "earnings.xlsx",
  "mode": "content",
  "edits": [
    { "id": "sheet:Summary/cell:B2", "text": "Q4 FY2025" },
    {
      "id": "sheet:Summary/chart:1/title",
      "text": "Revenue vs EBITDA by Quarter"
    }
  ]
}
```

## Preview edits

Tool: `workspace_preview_edits`

```json
{
  "filename": "earnings.xlsx",
  "mode": "content",
  "lint": true,
  "strict": true,
  "edits": [{ "id": "sheet:Summary/cell:B2", "text": "Q4 FY2025" }]
}
```

## Apply edits

Tool: `workspace_apply_edits`

```json
{
  "filename": "earnings.xlsx",
  "mode": "content",
  "lint": true,
  "strict": true,
  "output_filename": "earnings.updated.xlsx",
  "edits": [{ "id": "sheet:Summary/cell:B2", "text": "Q4 FY2025" }]
}
```
