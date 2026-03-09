# Interop Feature Buckets

Status: Draft
Last updated: 2026-03-05
Related: `INTEROP.md`, `EDITOR_PARITY.md`

This file keeps interoperability scope explicit.

Every imported/exported feature should live in one bucket:

- `editable`
- `preserve-only`
- `warn-and-drop`

## Editable

These are the common document features the editor should actively support.

- plain text
- paragraphs and line breaks
- headings
- bold / italic / underline / strikethrough
- font family / size / text color
- paragraph alignment
- basic lists and numbering
- basic indentation and spacing
- simple tables
- inline images
- links
- basic section/page properties if stable:
  - page size
  - margins
  - orientation

## Preserve-Only

These should survive open/save when possible, but direct editing is not yet the
goal.

- custom styles and theme mappings
- advanced table layout:
  - merges
  - repeating headers
  - exact width behavior
- floating images / anchored objects
- headers / footers beyond basic text
- footnotes / endnotes
- fields:
  - TOC
  - PAGE / NUMPAGES
  - references
- comments
- track changes / revisions
- bookmarks and cross-references
- content controls
- advanced numbering definitions
- tabs / leaders with complex behavior
- DrawingML effects and shapes
- SmartArt / charts / embedded objects
- compatibility settings and obscure OOXML flags

## Warn-And-Drop

This bucket should stay small and explicit.

- active content or unsafe embedded payloads
- unsupported binary embeddings with no safe preservation path
- corrupt or invalid OOXML fragments we must sanitize
- vendor-specific extensions that break save/export invariants

## Rules

- `editable`: user can modify directly and expect save/open to retain intent
- `preserve-only`: user may see it, but editing is not guaranteed; save should keep it if untouched
- `warn-and-drop`: only use when preservation is unsafe or impossible

## Current 80/20 Priority

1. keep common authoring features `editable`
2. keep complex Word-only structures `preserve-only`
3. keep `warn-and-drop` rare and visible
