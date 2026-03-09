# Editor Parity

Status: Draft
Last updated: 2026-03-05
Related: `PAGINATION_v2.md`, `EDITOR_PARITY_TRACKER.md`, `COLLABORATION.md`, `INTEROP.md`

This document defines the 80/20 product behavior for an internal
Word/Google-Docs-like editor built on top of the canvas layout engine.

## Scope

This spec is about editing behavior, not pagination.

It defines what users should expect when they type, move the caret, select
content, format text, and manipulate common document structures.

## Non-Goals

- full Microsoft Word compatibility
- full Google Docs feature parity
- editable tracked changes UI
- advanced floating object authoring
- exhaustive accessibility parity

## P0 Feature Set

- typing into body text and headings
- caret movement by keyboard and mouse
- multi-line selection
- delete and backspace
- paragraph split and merge
- copy, cut, paste
- undo, redo
- bold, italic, underline, strikethrough
- font family, font size, text color
- heading styles
- bulleted and numbered lists
- simple tables
- inline images

## Core Editing Rules

### Typing

- typing inserts at the caret
- typing replaces the active selection
- formatting at a collapsed selection becomes pending input formatting
- typing after a paragraph split should keep the expected inline formatting

### Enter And Backspace

- `Enter` splits the current paragraph
- `Shift+Enter` inserts a soft line break
- `Backspace` joins with the previous paragraph when at paragraph start
- `Delete` joins with the next paragraph when at paragraph end
- split and join behavior must preserve intuitive formatting

### Selection

- support collapsed, forward, and backward selection
- support mouse drag, shift-click, and keyboard extension
- double click selects a word
- triple click selects a paragraph
- selections across paragraphs must delete, copy, and replace correctly

### Keyboard Navigation

- arrows move by character and line
- modified arrows move by word when supported by platform conventions
- `Home` and `End` move to line boundaries
- `Cmd/Ctrl+A` selects the whole document
- `Cmd/Ctrl+Z` and `Cmd/Ctrl+Shift+Z` or `Ctrl+Y` undo and redo
- `Cmd/Ctrl+B`, `I`, `U` toggle common formatting

## Formatting Model

### Inline Formatting

- bold
- italic
- underline
- strikethrough
- font family
- font size
- text color

### Block Formatting

- body text
- heading 1
- heading 2
- heading 3
- paragraph alignment
- list on/off
- list indent and outdent

### Mixed Selections

- toolbar state must handle mixed formatting explicitly
- applying a format to a mixed selection should normalize the selected range

## Lists

- support bullet and numbered lists
- `Enter` creates a new item
- `Enter` on an empty item exits the list
- `Tab` and `Shift+Tab` indent and outdent
- numbering continuation must remain stable after edits

## Tables

The first table target is simple and practical:

- insert a rectangular table
- place caret inside cells
- navigate cells with keyboard
- add/remove rows and columns
- delete content inside cells

Advanced merge/split, resize, and complex table styling can come later.

## Images

The first image target is also simple:

- insert inline image
- select image
- delete image
- preserve size and placement on save

Resize and floating authoring are later features unless explicitly promoted.

## IME And Composition

- hidden input host is authoritative for composition
- composition must work for CJK input, dead keys, and emoji
- caret and selection must remain stable during composition
- remote edits should not corrupt in-progress composition state

## Clipboard

- copy and cut preserve selected text and common formatting
- paste replaces active selection
- prefer rich HTML plus plain text when available
- pasted lists, tables, and images should degrade predictably, not randomly

## Undo And Redo

- sequential typing should coalesce into sensible undo groups
- paste is one undo unit
- formatting operations are undoable
- table and image insertions are undoable
- remote collaboration must not corrupt local undo state

## Deferred

- comments authoring
- tracked changes UI
- advanced table authoring
- advanced image wrapping/crop
- full field editing

## Acceptance Bar

This spec is met when common authoring flows feel boring in the best way:

- the caret goes where users expect
- selections delete what users think is selected
- formatting sticks
- keyboard shortcuts behave predictably
- lists, simple tables, and inline images are usable end to end
