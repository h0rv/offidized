# Editor Parity Tracker

Status: Draft
Last updated: 2026-03-06
Related: `EDITOR_PARITY.md`, `PAGINATION_v2.md`, `COLLABORATION.md`

Legend: `todo`, `in_progress`, `done`, `deferred`, `blocked`

This tracker is intentionally lean. It exists to keep the editor work in scope
and current, not to become a second roadmap.

## Core Editing

Status: `done`

- [x] insert text at caret / replace selection
- [x] `Enter` splits paragraph
- [x] `Shift+Enter` inserts line break
- [x] `Backspace` and `Delete` inside paragraph
- [x] backspace at paragraph start merges backward
- [x] delete at paragraph end merges forward
- [x] paste replaces selection
- [x] cut/delete selection

## Selection And Caret

Status: `done`

- [x] click places caret correctly
- [x] drag selection works across lines
- [x] backward selection behaves correctly
- [x] double-click selects word
- [x] triple-click selects paragraph
- [x] selection highlight matches deleted content
- [x] caret geometry stays stable while typing and moving
- [x] caret/selection works across paragraph boundaries

Notes:

- the canvas/controller path now treats double click as word selection and triple click as full-paragraph selection
- backward selection restore now has browser-level HTML acceptance, and the HTML rect helper now normalizes backward anchor/focus before building geometry
- browser ops acceptance now covers exact drag-range deletion, reverse drag copy/delete, and broader cross-paragraph HTML selection behavior
- cross-paragraph keyboard selection/collapse also has focused controller coverage
- click-to-caret placement now has focused controller coverage, and the browser caret harness is green again

## Keyboard Navigation

Status: `done`

- [x] arrow keys
- [x] `Home` / `End`
- [x] `PageUp` / `PageDown`
- [x] `Shift+Arrow` extends selection
- [x] `Cmd/Ctrl+A`
- [x] word navigation with modifiers
- [x] word delete with modifiers

Notes:

- the controller path now has focused acceptance for left/right character movement, cross-paragraph backward movement, range-collapse semantics, `Home` / `End`, and hit-test-driven `PageUp` / `PageDown`
- cross-paragraph `Shift+ArrowLeft` / `Shift+ArrowRight` extension and backward multi-paragraph collapse are also covered now
- focused controller acceptance also exists for select-all plus modifier word movement and word deletion

## IME And Composition

Status: `done`

- [x] composition start/update/commit
- [x] CJK typing works
- [x] emoji and surrogate pair safety
- [x] dead key / accent input
- [x] caret stays stable during composition
- [x] undo works after composition

Notes:

- canvas and HTML input now normalize committed IME text through a first-class composition commit path instead of flattening it into plain `insertText`
- focused IME tests now cover Japanese/Chinese committed composition text, dead-key / `Process` keyboard noise, and single-commit fallback behavior in both input adapters
- committed composition text now forms its own undo boundary, and focused controller coverage proves undo/redo behavior for that path
- the hidden canvas textarea now follows the visual caret overlay and keeps the last anchor while the overlay is temporarily hidden during composition

## Clipboard

Status: `done`

- [x] copy plain text
- [x] copy rich text
- [x] cut
- [x] paste plain text
- [x] paste rich text from Word / Docs
- [x] multi-paragraph paste
- [x] paste into selection

Notes:

- clipboard now ships the real 80/20 split: `text/plain` copy/cut/paste semantics plus `text/html` copy/cut/paste for the formatting and paragraph structure we model cleanly
- rich HTML paste currently covers paragraph blocks plus inline bold/italic/underline/strikethrough/font/color/highlight/link styling, heading/alignment cues, and straightforward `<ul>/<ol>/<li>` list structure
- accepted coverage now includes controller tests, save/reload coverage for rich pasted content, and browser ops acceptance in both HTML and canvas renderers

## Formatting

Status: `in_progress`

- [x] bold
- [x] italic
- [x] underline
- [x] strikethrough
- [x] font family
- [x] font size
- [x] text color
- [x] text highlight
- [x] link / hyperlink
- [x] collapsed-selection pending formatting
- [x] mixed-selection toolbar state
- [x] save/reload preserves formatting

Notes:

- core inline formatting now ships through the text-attr path with controller authoring, formatting-state readback, save/reload roundtrip, and basic diff-based peer convergence
- hyperlink authoring now ships through the text-attr path, with save/reload roundtrip and basic diff-based peer convergence
- text highlight now ships through the same text-attr path, with toolbar state readback, save/reload roundtrip, and basic diff-based peer convergence
- collapsed-selection formatting is now treated as pending input formatting in the controller path and applies to newly typed text
- mixed inline and paragraph selections now report non-uniform toolbar state honestly instead of falsely claiming a single bold/alignment value

## Paragraphs And Headings

Status: `in_progress`

- [x] body text
- [x] heading 1
- [x] heading 2
- [x] heading 3
- [x] left alignment
- [x] center alignment
- [x] right alignment
- [x] justify alignment
- [x] space before paragraph
- [x] space after paragraph
- [x] line spacing presets
- [x] left indent
- [x] first-line indent
- [x] toggle heading on existing paragraph
- [x] save/reload preserves heading style
- [x] save/reload preserves alignment
- [x] save/reload preserves basic paragraph spacing/indent
- [x] save/reload preserves line spacing presets

Notes:

- paragraph alignment now ships through the actual editor path: toolbar/controller authoring, formatting-state readback, save/reload roundtrip, and basic diff-based peer sync
- the first paragraph spacing/indent slice now also ships through the actual editor path: space before, space after, left indent, and first-line indent
- line spacing now ships as a deliberately narrow Docs-style preset slice on the paragraph-attr path: `1.0`, `1.15`, `1.5`, and `2.0` map to Word `auto` spacing and have formatting-state readback, save/reload roundtrip, and basic diff-based peer sync
- minimal alignment acceptance should stay narrow: authoring via paragraph attrs, roundtrip through save/reload, and basic peer convergence

## Lists

Status: `in_progress`

- [x] imported bullet / numbering metadata reaches the editor model
- [x] editor view model resolves list prefix text when the original document numbering definitions are available
- [x] export preserves numbering metadata when list attrs are present in CRDT paragraph state
- [x] toggle list on selected paragraphs
- [x] `Enter` creates next list item
- [x] `Enter` on empty item exits list
- [x] load/edit/save roundtrip for existing list paragraphs is covered end to end
- [x] `Tab` / `Shift+Tab` indent and outdent
- [x] numbering continues correctly after in-editor edits

Notes:

- current branch stores `numberingKind`, `numberingNumId`, and `numberingIlvl` in editor paragraph state and has explicit numbering-aware export logic
- the controller path now supports bullet/decimal toggles plus list continuation and empty-item exit on `Enter`
- the controller path also treats `Tab` / `Shift+Tab` as list indent/outdent, while keeping plain-paragraph `Tab` as a literal tab token
- controller-authored list lifecycle operations now have focused peer-convergence coverage for toggle, continue, empty-item exit, and indent/outdent
- save/reload now writes the missing numbering package parts, so reopened editors can recover list kind and numbering text instead of losing list state
- list markers are rendered visually without counting as editable text, which keeps caret and selection offsets aligned
- controller numbering now reattaches to adjacent list segments instead of needlessly forking `numId`s, and focused tests cover renumbering after deletion plus middle-paragraph reattachment

## Tables

Status: `in_progress`

- [x] insert basic table
- [x] edit plain text in a cell through the overlay editor
- [x] `Tab` moves cell-to-cell
- [x] add/remove row
- [x] add/remove column
- [x] delete inside cell
- [x] copy/paste text into table
- [x] save/reload preserves table edits

Notes:

- the shipped controller path inserts a basic table, immediately opens the first cell in a DOM overlay editor, commits text back through `setTableCellText`, and advances with `Tab`
- focused controller coverage now saves/reloads edited cell text, peer-syncs those edits to a second editor, and the browser autorun harness covers toolbar insert plus `Tab` navigation
- row/column structure edits now have controller, save/reload, peer-sync, and browser autorun coverage
- the table cell textarea now commits live on `input`, `cut`, and `paste`, so delete/backspace plus plain-text clipboard edits persist immediately instead of waiting for blur
- accepted coverage for table delete/copy/paste is the real browser autorun path against the overlay textarea in both renderers
- richer table clipboard behavior and non-text cell content remain separate work

## Images

Status: `in_progress`

- [x] insert inline image
- [x] select image
- [x] delete image
- [x] paste image
- [x] resize selected inline image in the editor
- [x] left / center / right block alignment for selected inline image paragraphs
- [x] save/reload preserves image

Notes:

- the shipped demo/controller path now reads an image file from the toolbar, derives a size from the decoded asset, and inserts it through `insertInlineImage`
- inline image deletion now works on the token path through normal backspace/delete semantics
- focused controller coverage now saves/reloads inline image runs, resizes selected images, and syncs insert/delete plus resize/alignment to a peer
- explicit inline-image click selection and clipboard image paste now ship in both renderers
- the current 80/20 resize UI is a selected-image overlay with a single drag handle; it mutates the existing inline-image token attrs instead of introducing floating-image editing
- block image alignment reuses paragraph alignment for image paragraphs; floating positioning and wrap editing remain deferred

## Undo / Redo

Status: `in_progress`

- [x] undo typing
- [x] redo typing
- [x] undo formatting
- [x] undo paste
- [x] undo table/image insert
- [x] history grouping feels saner for typed replacement and semantic boundaries
- [x] undo remains local under collab

Notes:

- redundant/no-op remote updates no longer clear local undo history
- real remote mutations now preserve prior local undo history by expanding the existing undo scope instead of rebuilding the manager

## Collaboration Basics

Status: `in_progress`

- [x] remote text sync
- [x] remote formatting sync
- [x] remote selection/cursor presence
- [x] reconnect/resync after desync
- [x] copy-link-and-join works
- [x] basic conflict convergence

Notes:

- browser-level acceptance now proves a copied collaboration URL preserves the active `room` and `ws` params needed for another tab to join the same session

## Deferred

Status: `deferred`

- [ ] comments
- [ ] track changes / suggestions
- [ ] advanced table features
- [ ] floating images / wrap editing
- [ ] footnotes/endnotes authoring
- [ ] advanced accessibility parity

## Current Priorities

No active non-deferred parity gaps remain in this tracker. Remaining work is intentionally deferred.
