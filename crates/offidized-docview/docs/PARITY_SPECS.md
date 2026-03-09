# Editor Parity Spec Set

Status: Draft
Last updated: 2026-03-05
Audience: offidized-docview contributors

This directory now has a small spec set that defines the practical
"Word/Docs-like editor" target:

| Spec                         | Purpose                                                             |
| ---------------------------- | ------------------------------------------------------------------- |
| `PAGINATION_v2.md`           | Deterministic layout, pagination, hit-testing, and canvas rendering |
| `EDITOR_PARITY.md`           | Editing semantics and user-visible behavior                         |
| `EDITOR_PARITY_TRACKER.md`   | Lean implementation/status checklist for editor behavior            |
| `COLLABORATION.md`           | Shared editing, presence, convergence, resync                       |
| `COLLABORATION_CHECKLIST.md` | Lean implementation/status checklist for collaboration              |
| `INTEROP.md`                 | `.docx` import/export fidelity and preserve-vs-edit policy          |
| `INTEROP_FEATURE_BUCKETS.md` | Explicit `editable` vs `preserve-only` vs `warn-and-drop` buckets   |

## Intent

The goal is not full Microsoft Word feature parity on day one.

The goal is an internal editor that feels familiar, stable, and useful for the
common 80/20 workflows:

- typing and navigation
- selection and deletion
- copy/cut/paste
- undo/redo
- headings and common formatting
- lists
- simple tables
- inline images
- live collaboration
- `.docx` open and save without destructive loss

## Explicitly Out Of Scope For This Spec Set

- rollout strategy
- production ops and QA process
- spreadsheet-style product scope
- enterprise permissions/auth

## Priority Order

1. `PAGINATION_v2.md`
   Layout must be deterministic enough for caret, selection, and document trust.

2. `EDITOR_PARITY.md`
   If typing, selection, deletion, and formatting are off, the editor is not
   usable even if pagination is good.

3. `COLLABORATION.md`
   Shared editing must converge and recover cleanly from transient desync.

4. `INTEROP.md`
   Open/save must preserve the common document shapes users actually care about.

## Feature Classification

### P0: Must Feel Solid

- text input
- caret movement
- selection
- delete/backspace
- enter/shift-enter
- copy/cut/paste
- undo/redo
- bold/italic/underline
- headings
- font family / size / color
- lists
- simple tables
- inline images
- shared editing with remote cursors
- `.docx` open/save for common docs

### P1: Strongly Desirable

- richer paragraph styling
- table row/column controls
- image resize
- comments preservation
- better style fidelity
- footnotes rendering and preservation

### Later

- full tracked changes UI
- full comments workflow
- advanced floating objects
- deep field evaluation
- full Word compatibility-flag matrix

## Working Rule

When there is tension between architectural cleverness and predictable user
behavior, prefer the simpler behavior that matches Word/Docs expectations.

## Keep-It-Lean Rule

These docs should stay small and current:

- specs define behavior and boundaries
- trackers define status and immediate priorities
- feature buckets define what is editable versus preserved

If a feature lands, update the relevant tracker or bucket in the same change.
