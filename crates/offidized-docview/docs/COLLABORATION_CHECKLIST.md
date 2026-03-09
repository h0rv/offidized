# Collaboration Checklist

Status: Draft
Last updated: 2026-03-06
Related: `COLLABORATION.md`, `EDITOR_PARITY.md`

Legend: `todo`, `in_progress`, `done`, `deferred`, `blocked`

This is the lean execution checklist for shared editing.

## Shared State

Status: `done`

- [x] room/document id model is stable
- [x] multiple peers attach to same room
- [x] local edits become shared CRDT updates
- [x] peers converge after normal editing

## Presence

Status: `in_progress`

- [x] relay has a separate ephemeral awareness transport path (`0x01` WebSocket frames, not stored)
- [x] renderer adapters expose cursor/selection geometry hooks for peer overlays
- [x] remote cursor
- [x] remote selection
- [x] participant label/color
- [x] stale presence expires automatically
- [x] BroadcastChannel local-awareness fallback is wired and accepted
- [x] explicit peer-drop cleanup converges without stale peer resurrection

## Join And Resync

Status: `done`

- [x] new peer can join existing room
- [x] new peer gets current state before editing
- [x] reconnect requests missing updates
- [x] fallback full-state resync works
- [x] transport state is visible in UI

Notes:

- the browser collab harness now proves a stronger sequence: rich late join into an already-edited room, live replace convergence, and same-peer transport-pause repair after missed updates
- accepted reconnect coverage is the paused-peer repair path, not iframe recreation
- the stale peer explicitly requests current state on resume and converges on the latest text/table/image state

## Desync Handling

Status: `done`

- [x] detect likely divergence
- [x] trigger anti-entropy repair
- [x] avoid silent long-lived desync
- [x] room stays usable after repair

Notes:

- divergence visibility is now accepted on the browser harness path through a paused-peer repair flow
- the stale peer visibly enters repair/desync, emits repair counters, requests current state, and clears those flags after convergence
- the healthy responder no longer falsely marks itself as desynced just for serving current state to a stale peer

## Formatting And Rich Content

Status: `in_progress`

- [x] formatting changes sync
- [x] headings sync
- [x] paragraph alignment sync
- [x] basic paragraph spacing/indent sync
- [x] line spacing preset sync
- [x] highlight attrs sync
- [x] hyperlink attrs sync
- [x] existing list metadata convergence is covered end to end
- [x] list authoring operations sync
- [x] table edits sync
- [x] image insert/delete sync
- [x] image resize/alignment sync

Notes:

- paragraph alignment now has the minimum accepted coverage: authoring through paragraph attrs, save/reload roundtrip, and basic diff-based peer convergence
- the first paragraph spacing/indent slice now has the same minimum accepted coverage on the paragraph-attr path: authoring, save/reload, and basic diff-based peer convergence
- line spacing presets now have the same minimum accepted coverage on the paragraph-attr path, intentionally scoped to Word `auto` spacing multiples rather than `exact` / `atLeast`
- highlight attrs now have the same minimum accepted coverage on the text-style path: authoring, save/reload, and basic diff-based peer convergence
- hyperlink attrs now have the same minimum accepted coverage on the text-style path: authoring, save/reload, and basic diff-based peer convergence
- paragraph sync now has the raw shape needed for lists because paragraph attrs can carry `numberingKind`, `numberingNumId`, and `numberingIlvl`
- focused CRDT peer coverage now exists for an existing list paragraph, including numbering metadata plus text formatting convergence
- that supports the current checklist item for existing-list collaboration on the existing paragraph-attribute path
- controller-authored list lifecycle operations now have focused peer coverage: toggle, continue on `Enter`, empty-item exit, and `Tab` / `Shift+Tab` indent-outdent
- broader browser-level list UX and numbering behavior after arbitrary edits remain separate work
- focused peer-convergence coverage now exists for controller-authored table cell edits and inline-image insert/delete/resize/alignment
- the browser collab harness now proves those richer states on the real multi-window path in addition to text and presence convergence
- richer structural ops now force an immediate full-state flush in addition to diff sync so browser peers converge on table/image changes without waiting for the periodic anti-entropy timer
- presence transport is intentionally ephemeral and separate from CRDT state; current relay behavior broadcasts awareness to connected peers only and does not replay it to late joiners
- peer overlay geometry is available in both renderers and the browser collab harness now verifies peer overlay presence in addition to text convergence
- explicit peer-drop presence cleanup is now accepted on the real awareness path; the surviving peer keeps a sender tombstone so late queued awareness packets cannot resurrect stale overlays
- automatic stale peer expiry is now accepted on the browser-proven timer path; the collab harness asserts expiry without relying on explicit `bye`

## Undo Expectations

Status: `todo`

- [x] redundant/no-op remote updates do not clear local undo
- [x] local undo stays coherent under real remote edits
- [x] remote edits do not corrupt local history

## Save And Replace

Status: `done`

- [x] save exports converged local state
- [x] opening a new doc replaces room state cleanly
- [x] peers see replace event consistently

Notes:

- room-level replace is generation-aware, so opening a new blank/docx snapshot replaces peer state instead of merging into the old document
- the browser collab harness now proves replace propagation clears prior table/image state and converges on the new text snapshot

## Deferred

Status: `deferred`

- [ ] auth and permissions
- [ ] comments collaboration
- [ ] suggestions / tracked changes collaboration
- [ ] role-based editing

## Current Priorities

1. richer table/image authoring beyond the current 80/20 slice
2. richer clipboard/import paths beyond plain text
