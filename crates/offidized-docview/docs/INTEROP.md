# Interoperability

Status: Draft
Last updated: 2026-03-05
Related: `PAGINATION_v2.md`, `EDITOR_PARITY.md`, `COLLABORATION.md`, `INTEROP_FEATURE_BUCKETS.md`

This document defines the 80/20 `.docx` interoperability contract.

## Scope

The target is practical document fidelity:

- open normal `.docx` files cleanly
- edit them in the internal editor
- save valid `.docx` files that reopen correctly
- avoid destructive loss for unsupported structures

## Non-Goals

- full OOXML editing parity
- full track changes workflow
- full comments workflow
- full field evaluation engine

## Canonical Rule

- the editor model is the live editing model
- `.docx` is the import/export boundary
- preserve-first is preferred over normalize-and-destroy

## Open Semantics

- import text, paragraphs, headings, styles, lists, tables, images, and sections commonly used in normal docs
- preserve unsupported structures when possible
- if a structure cannot be edited faithfully, prefer preserve-only behavior

## Save Semantics

- export current editor state back to valid `.docx`
- keep unchanged original structures intact when possible via clone-and-patch
- do not silently drop known important content

## P0 Fidelity Targets

- paragraph boundaries
- inline formatting
- font family, size, color
- headings and basic paragraph styles
- lists
- simple tables
- inline images
- common sections and page settings

## Style And Theme Policy

- preserve style ids and inheritance when possible
- allow editing of common style effects without requiring full style-ui parity
- preserve theme references when not directly edited

## Fields

- preserve field codes and results where possible
- do not require full live field evaluation in the first pass
- complex fields are preserve-first unless explicitly supported

## Comments And Track Changes

- preserve on roundtrip when possible
- editing UI is deferred
- behavior around editing inside revision-heavy regions must be defined before broad use

## Tables And Images

- support simple rectangular tables for import, edit, and save
- preserve richer table structures where possible
- support inline image import and save first
- floating image fidelity is later unless explicitly promoted

## Unsupported Feature Policy

Every feature should fall into one of three buckets:

- editable
- preserve-only
- dropped with explicit warning

Silent loss is not acceptable.

## Acceptance Bar

This spec is met when:

- common Word-authored documents open and save cleanly
- minimal edits do not destroy unrelated structure
- unsupported content is preserved whenever practical
- saved files reopen in Word without obvious breakage
