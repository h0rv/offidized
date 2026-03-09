# Collaboration

Status: Draft
Last updated: 2026-03-06
Related: `PAGINATION_v2.md`, `EDITOR_PARITY.md`, `COLLABORATION_CHECKLIST.md`, `INTEROP.md`

This document defines the 80/20 collaboration model for the editor.

## Scope

The collaboration target is simple:

- two or more people can edit the same document
- content converges
- users can see each other
- transient disconnects do not permanently break the session

## Non-Goals

- auth and permissions
- enterprise audit history
- comments and suggestion workflows
- complex conflict-resolution UI

## Core Model

- the CRDT document is the only mutable source of truth
- layout is derived from CRDT state and is never the sync authority
- local edits become intents, then CRDT transactions, then shared updates

## Session Model

- each collaborative document has a stable room/document id
- multiple clients can attach to the same room
- a new client must receive enough state to converge before being considered ready

## Presence

Presence is ephemeral and separate from document content.

Presence transport currently uses a dedicated awareness lane, separate from CRDT
document updates:

- WebSocket relay message type `0x01` is reserved for awareness fan-out in [demo/relay-server.ts](../demo/relay-server.ts)
- local multi-window fallback also carries awareness over `BroadcastChannel`
- awareness is broadcast to currently connected peers only and is not stored or replayed
- CRDT content sync still uses the normal update path; presence is not part of document save/export

Current presence payload should stay minimal:

- color
- caret position
- selection range
- optional short peer label

Current implementation note:

- renderer adapters now expose document-position-to-rect helpers for peer overlays in both HTML and canvas renderers
- the editor controller renders remote presence as a non-document overlay; peer cursor/selection visuals never mutate CRDT content
- the current shipped overlay renders remote selection blocks, a remote caret bar, and a short peer label
- explicit peer-drop cleanup is now covered on the real awareness path, including sender tombstones so queued stale awareness packets cannot resurrect a disconnected peer
- automatic stale expiry is now also accepted on the browser harness path; expiry is timer-driven and does not rely on explicit `bye` delivery
- local undo now survives real remote edits by keeping the existing undo manager alive across non-noop remote updates and expanding its scope for newly introduced remote paragraphs
- browser-level acceptance now covers rich late join into an already-edited room, live replace propagation, and same-peer repair after missed updates via a paused-transport harness path

## Remote Cursors And Selections

- peer cursors and selections should render as overlays derived from document positions, not stored pixel coordinates
- the overlay geometry now comes from renderer-level rect queries rather than ad hoc DOM scraping
- awareness is ephemeral; losing a peer cursor must never mutate document state
- late joiners should not expect historical cursor/selection replay because awareness is not persisted

## Convergence

- all peers must converge to the same CRDT state after updates are exchanged
- transport reordering must not change final content
- rendering lag is acceptable; document divergence is not

Current list-status note:

- the branch now carries list-related paragraph attrs in CRDT state (`numberingKind`, `numberingNumId`, `numberingIlvl`)
- focused peer-edit coverage now exists for an existing list paragraph: list numbering metadata and text formatting converge across CRDT peers
- controller-authored list lifecycle operations now also have focused peer coverage: toggle, continue on `Enter`, empty-item exit, and `Tab` / `Shift+Tab` indent-outdent
- controller-side list ID reuse now reattaches adjacent decimal segments instead of needlessly forking numbering when a middle paragraph is toggled back into the list
- this supports an honest claim that list authoring collaboration is working for the current paragraph-attribute path
- broader browser-level list UX remains separate work

Current rich-content note:

- local authoring now includes basic table insert with plain cell-text edits plus inline image insert/delete
- focused peer convergence coverage now exists for controller-authored table cell edits and inline image insert/delete/resize/alignment
- the browser collaboration harness also proves table cell edits and inline image insert converge in a real two-window session

## Reconnect And Resync

- on reconnect, the client rejoins the room
- the client advertises its current known state
- the peer or server provides missing updates or a full snapshot
- if incremental sync confidence is low, force a full resync

Presence note:

- document resync and presence recovery are separate
- CRDT content uses explicit anti-entropy repair
- explicit peer-drop cleanup and timer-driven stale expiry are both accepted
- accepted reconnect coverage now uses a stale peer that pauses transport, misses updates, resumes, requests current state, and converges without recreating the client instance

## Desync Handling

When the system detects likely divergence:

- mark the session as resyncing
- request authoritative state
- repair automatically if possible
- avoid lying to the user that sync is healthy when it is not

Current implementation note:

- divergence visibility is latched on the stale peer during the browser harness repair path, so the acceptance signal does not depend on catching a transient final-state flag
- healthy responders no longer enter repair/desync state merely because another peer asked for current state

## Save And Replace

- save exports the local converged CRDT state to `.docx`
- save is not a collaboration checkpoint
- opening a new `.docx` is a full document replacement event for the room

## Undo Expectations

- undo is local to the user
- local undo must stay coherent under remote changes
- if a behavior cannot be made intuitive, it should be deferred rather than guessed

Current implementation note:

- redundant/no-op remote updates keep the current local undo history intact
- real remote mutations now preserve prior local undo history by expanding the existing undo scope rather than rebuilding the undo manager

## Deferred

- comments presence
- suggestions / tracked changes collaboration
- permissions
- role-based editing

Current presence limitations:

- presence is not stored, replayed, or exported
- viewport/page presence is not implemented
- participant labels/colors remain minimal and are intended only to disambiguate peers, not provide a full roster model

## Acceptance Bar

This spec is met when:

- two tabs converge under normal typing, deletion, paste, and formatting
- existing list paragraphs also converge under peer edits without losing numbering metadata
- basic table cell edits and inline image insert/delete also converge under peer edits
- remote cursors and selections remain legible in both renderer modes
- reconnect repairs transient drift
- the editor does not silently remain in a desynced state
