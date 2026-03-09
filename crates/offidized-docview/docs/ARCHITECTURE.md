# offidized-docview: Collaborative Editor Architecture

This document describes the layers required to evolve offidized-docview from a
read-only viewer into a full collaborative document editor backed by CRDTs.

## Design Principle: Single Source of Truth

The CRDT document is the sole mutable source of truth during editing.
There are not three mutable models that need reconciliation — there is one:

```
         ┌──────────────┐
  .docx ─┤  IMPORT      │
  bytes  │  (one-time)  ├──► CRDT Doc ◄──── all edits happen here
         └──────────────┘    │       │
                             │       │ on save / on snapshot
                  (retained) │       ▼
         ┌──────────────┐    │  ┌──────────────┐
         │  Original    │◄───┘  │  EXPORT      │──► .docx bytes
         │  Package     │──────►│  (clone +    │
         │  (immutable) │       │   patch)     │
         └──────────────┘       └──────────────┘
```

- **offidized-docx** is used for import (parse .docx → populate CRDT) and
  export (clone original package → patch only changed parts). It is never
  mutated during the editing loop.
- **The CRDT** owns all live document state. Edits go directly to it.
  The view model is derived from the CRDT, not from offidized-docx.
- **No intermediate DocOp layer** that mutates a separate model. User
  intents map directly to CRDT transactions.
- **The original parsed Document is retained immutably.** On export, it
  is cloned and only the parts that differ from CRDT state are patched.
  Styles, themes, and non-editable OOXML parts survive untouched on
  the clone. Paragraph-level and above unknown elements (`RawXmlNode`)
  are preserved in their original locations because only runs inside
  dirty paragraphs are replaced. Run-level unknown elements are
  preserved via opaque tokens in the CRDT (extracted on import,
  re-emitted on export).

## Current State

What exists today:

- **Rust (WASM)**: `DocView` struct parses `.docx` bytes via `offidized-docx`,
  converts to `DocViewModel` (JSON-serializable, CSS-point measurements),
  returns it to JS.
- **TypeScript**: `DocRenderer` consumes `DocViewModel`, builds DOM.
  `mount()` API and `<doc-view>` web component. Paginated and continuous
  modes. No editing, no mutation path.
- **offidized-docx**: Full mutable document model — paragraphs, runs,
  tables, images, footnotes, comments, numbering, styles, sections.
  Roundtrip fidelity via `RawXmlNode` preservation.

---

## Target Architecture

```
┌──────────────────────────────────────────────────────────────┐
│ Browser                                                      │
│                                                              │
│  ┌────────────────────────────────────────────────────────┐  │
│  │ TypeScript: Editor Shell                               │  │
│  │ contenteditable, Selection API, IME, clipboard         │  │
│  │ Keyboard/mouse → editing intents                       │  │
│  │ Remote cursor rendering                                │  │
│  │ DOM patching from view updates                         │  │
│  └──────────────────────┬─────────────────────────────────┘  │
│                         │ intents + CRDT-relative positions   │
│  ┌──────────────────────▼─────────────────────────────────┐  │
│  │ Rust/WASM: offidized-docview                           │  │
│  │                                                        │  │
│  │  ┌──────────────────────────────────────────────────┐  │  │
│  │  │ CRDT Doc (single source of truth)                │  │  │
│  │  │ Rich text sequences with formatting attributes   │  │  │
│  │  │ Awareness (cursors, user presence)               │  │  │
│  │  │ Undo manager                                     │  │  │
│  │  │ Sync encoding (state vectors, incremental diffs) │  │  │
│  │  └──────────┬────────────────────┬──────────────────┘  │  │
│  │             │                    │                      │  │
│  │  ┌──────────▼──────────┐  ┌─────▼──────────────────┐  │  │
│  │  │ View Model          │  │ Import / Export         │  │  │
│  │  │ CRDT → DocViewModel │  │ .docx ↔ CRDT           │  │  │
│  │  │ CRDT change events  │  │ Uses offidized-docx    │  │  │
│  │  │ → ViewPatch list    │  │ for parse/serialize    │  │  │
│  │  └─────────────────────┘  └────────────────────────┘  │  │
│  │                                                        │  │
│  └────────────────────────────────────────────────────────┘  │
│                         │                                    │
│  ┌──────────────────────▼─────────────────────────────────┐  │
│  │ TypeScript: Persistence Providers                      │  │
│  │ All storage/network lives in TS (browser APIs)         │  │
│  │ WASM exposes: encode_state_vector(), encode_update(),  │  │
│  │   apply_update(), snapshot_to_docx()                   │  │
│  └────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────────────┐
│ Collaboration Server                                         │
│ WebSocket relay, awareness broadcast, update persistence     │
│ Auth, room management, .docx snapshot export                 │
└──────────────────────────────────────────────────────────────┘
```

## Feature Flags

```toml
[features]
default = []
editing = ["<crdt-library>"]       # CRDT is the editing model from day 1
sync = ["editing"]                  # state vector encoding, incremental diffs
awareness = ["sync"]                # cursor sharing, user presence
```

Three npm packages:

| Package              | Features                       | Use case              |
| -------------------- | ------------------------------ | --------------------- |
| `@offidized/docview` | none                           | Read-only viewer      |
| `@offidized/docedit` | `editing`                      | Single-user editing   |
| `@offidized/doclive` | `editing`, `sync`, `awareness` | Collaborative editing |

The CRDT is the editing model from day 1 — even for single-user via
`@offidized/docedit`. It gives you undo/redo for free and avoids building
throwaway infrastructure. The `sync` and `awareness` features add the
incremental encoding, state vector diffing, and cursor sharing needed
for multi-client collaboration. `docedit` without `sync`/`awareness`
has a smaller WASM binary (no sync protocol code) and zero network
concepts — suitable for embedded editors that save via `.docx` export.

---

## CRDT as Document Model

### Structure

The CRDT document maps OOXML structure into CRDT-native types:

```
CrdtDoc
├── body: Array<Map>                    ← ordered body items
│   └── Map (paragraph)
│       ├── "id": string                ← app-level opaque ID (UUID/ULID)
│       ├── "type": "paragraph"
│       ├── "text": RichText            ← character sequence + formatting marks
│       ├── "alignment": string?
│       ├── "heading": u8?
│       ├── "numbering": Map?           ← { numId, ilvl }
│       ├── "style_id": string?
│       ├── "spacing_before": f64?
│       ├── "spacing_after": f64?
│       └── ...
│   └── Map (table)
│       ├── "id": string
│       ├── "type": "table"
│       ├── "rows": Array<Array<Map>>
│       │   └── Map (cell)
│       │       ├── "text": RichText
│       │       ├── "shading": string?
│       │       └── ...
│       └── ...
├── images: Map<string, Map>            ← content-hash → metadata (NOT binary)
├── footnotes: Array<Map>
├── endnotes: Array<Map>
└── styles: Map                         ← style registry (mostly read-only)
    (Paragraph-level+ RawXmlNodes, themes, doc properties live on the
     immutable original Document, NOT in the CRDT — see Export Strategy.
     Run-level RawXmlNodes are in the CRDT as opaque inline tokens.)
```

### Identity: App-Level Opaque IDs

Paragraphs and body items get app-level IDs (UUID or ULID), not CRDT
internal item IDs. The mapping between app IDs and CRDT positions is
maintained internally. This decouples identity from CRDT implementation
and survives library swaps.

```rust
/// App-level paragraph identifier. Opaque, stable across sessions.
/// Generated on import (new UUID per paragraph) or on creation.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct ParaId(pub [u8; 16]);  // UUID bytes

impl ParaId {
    pub fn new() -> Self { /* uuid v7 or ulid */ }
}
```

### Positions: CRDT-Relative, Not Flat Offsets

Cursor positions and edit targets use CRDT-native relative positions,
not `usize` flat-text offsets. This is critical because:

- Flat offsets break under concurrent inserts/deletes
- Flat offsets require expensive DOM text-length walking
- CRDT relative positions survive concurrent edits by definition

Each CRDT library has its own relative position API:

- **yrs**: `StickyIndex` / `Offset` relative to item IDs
- **automerge**: `Cursor` type (stable reference into a sequence)
- **loro**: `Cursor` with `Side::Left` / `Side::Right`

The WASM FFI exposes positions as opaque byte arrays:

```rust
/// Encode a CRDT-relative position from a paragraph ID + character index.
/// Returns an opaque blob that survives concurrent edits.
///
/// `char_index` is in UTF-16 code units (matching JS string indexing
/// and the DOM Selection API's offset semantics).
pub fn encode_position(para_id: &[u8], char_index: u32) -> Vec<u8>;

/// Decode a position back to (para_id, char_index) against current state.
/// Returns UTF-16 code unit offset.
pub fn decode_position(encoded: &[u8]) -> Option<(Vec<u8>, u32)>;
```

**Canonical position unit: UTF-16 code units.** This matches:

- JavaScript `String.length` and `String.prototype.charCodeAt()`
- DOM `Selection.anchorOffset` / `Selection.focusOffset` for text nodes
- `InputEvent.getTargetRanges()` offsets

The WASM side converts UTF-16 offsets to the CRDT's internal indexing
(which may be UTF-8, Unicode scalar values, or item IDs depending on
the library). The conversion happens inside `encode_position` /
`decode_position` — TypeScript never sees non-UTF-16 offsets.

This avoids cursor drift on emoji (surrogate pairs), combining characters,
and CJK IME composition sequences. The bridge always speaks UTF-16,
the CRDT always speaks its native unit, and the boundary is explicit.

TypeScript stores these opaque positions for selection state and passes
them back to WASM for edit operations.

### Rich Text ↔ Runs Mapping

CRDT rich text stores formatting as ranges/marks on a character sequence.
docx stores formatting as discrete `Run` objects with properties.

#### The inline token problem

docx runs can contain non-text inline content that doesn't reduce to
plain characters with formatting marks:

- **Field codes** (`w:fldChar` / `w:fldSimple`): multi-run constructs
  like PAGE, DATE, TOC entries. Consist of begin/separate/end markers
  spanning multiple runs.
- **Footnote/endnote references** (`w:footnoteReference`): single-run
  elements referencing a footnote by ID.
- **Bookmark markers** (`w:bookmarkStart` / `w:bookmarkEnd`): zero-width
  markers that span ranges.
- **Comment range markers** (`w:commentRangeStart` / `w:commentRangeEnd`).
- **Tab characters** (`w:tab`), **line breaks** (`w:br`).
- **Inline images** (`w:drawing` with `wp:inline`).
- **Unknown run children** (`RawXmlNode`): arbitrary XML that offidized-docx
  preserved but doesn't model.

Flattening these to plain text + formatting marks and coalescing back
on export will lose or misplace them.

#### Inline token model

Instead of pure text + marks, the CRDT rich text contains a mix of
**text characters** and **inline tokens**. Tokens are represented as
U+FFFC (Unicode Object Replacement Character) sentinel characters
with structured CRDT attributes. U+FFFC is the only sentinel — no
private-use-area characters are used (see "Sentinel character escaping"
below for the full escaping rules):

```
CRDT RichText for a paragraph:
  "Hello \uFFFC world \uFFFC."
         ^              ^
         │              └── token: { type: "footnoteRef", id: 1 }
         └── token: { type: "fieldSimple", fieldType: "PAGE",
                       instr: " PAGE ", presentation: "3" }

Formatting marks (orthogonal to tokens):
  [0..5]: { bold: true }
  [0..14]: { fontFamily: "Calibri" }
```

Each sentinel character carries a structured attribute map describing
what it represents. On export, the sentinel is expanded back into the
appropriate OOXML run content.

#### Token types

| Token type      | CRDT attribute                       | Export expansion                         |
| --------------- | ------------------------------------ | ---------------------------------------- |
| `fieldSimple`   | `{ fieldType, instr, presentation }` | `w:fldSimple` element                    |
| `fieldBegin`    | `{ id }`                             | `w:fldChar w:fldCharType="begin"` run    |
| `fieldCode`     | `{ id, instr }`                      | `w:instrText` run(s)                     |
| `fieldSeparate` | `{ id }`                             | `w:fldChar w:fldCharType="separate"` run |
| `fieldEnd`      | `{ id }`                             | `w:fldChar w:fldCharType="end"` run      |
| `footnoteRef`   | `{ id }`                             | `w:footnoteReference` run                |
| `endnoteRef`    | `{ id }`                             | `w:endnoteReference` run                 |
| `bookmarkStart` | `{ id, name }`                       | `w:bookmarkStart` element                |
| `bookmarkEnd`   | `{ id }`                             | `w:bookmarkEnd` element                  |
| `commentStart`  | `{ id }`                             | `w:commentRangeStart`                    |
| `commentEnd`    | `{ id }`                             | `w:commentRangeEnd`                      |
| `tab`           | `{}`                                 | `w:tab` element                          |
| `lineBreak`     | `{ type? }`                          | `w:br` element                           |
| `inlineImage`   | `{ imageRef, width, height }`        | `w:drawing` / `wp:inline`                |
| `opaque`        | `{ opaqueId, xml }`                  | Raw XML string re-emitted verbatim       |

#### Complex field codes: multi-token representation

Simple fields (`w:fldSimple`) map to a single `fieldSimple` token.
Complex fields (`w:fldChar` begin / instrText / separate / end) are
represented as **multiple sentinel tokens** that preserve the original
multi-run structure:

```
CRDT RichText for: "Page {PAGE field showing "3"} of document"

  "Page \uFFFC\uFFFC\uFFFC\uFFFC3\uFFFC of document"
        ^    ^    ^    ^     ^
        │    │    │    │     └── fieldEnd:      { id: "f1" }
        │    │    │    └── fieldSeparate:       { id: "f1" }
        │    │    └── fieldCode:               { id: "f1", instr: " PAGE " }
        │    └── (presentation text "3" is real text between separate/end)
        └── fieldBegin:                        { id: "f1" }
```

The `id` attribute groups the begin/code/separate/end tokens into one
logical field. Text between `fieldSeparate` and `fieldEnd` is the
presentation text — it's real CRDT text characters (editable, visible
to users as the field result). Text between `fieldBegin` and
`fieldSeparate` is the instruction string (not user-visible in the
editor, but preserved).

On export, each token maps 1:1 to its original OOXML run:

- `fieldBegin` → `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`
- `fieldCode` → `<w:r><w:instrText> PAGE </w:instrText></w:r>`
- `fieldSeparate` → `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`
- (presentation text) → normal text run(s)
- `fieldEnd` → `<w:r><w:fldChar w:fldCharType="end"/></w:r>`

This preserves the original multi-run structure exactly. No collapsing,
no lossy simplification. Nested fields (e.g., IF fields containing
PAGE fields) work naturally — each has its own `id`, and
begin/end nesting is tracked by ID during import and export.

The `opaque` token type is the escape hatch: any unknown run child that
offidized-docx preserved as `RawXmlNode` is serialized to an XML string
and stored as an opaque token attribute. On export, it's re-emitted as
raw XML. This guarantees no inline content is ever silently dropped.

Each opaque token carries an `opaqueId` (UUID v7, assigned on import).
This ID is stable across the token's lifetime and is used for
immutability enforcement: the validation layer tracks which `opaqueId`
values exist in the initial CRDT state and rejects any update that
mutates the `xml` or `opaqueId` attributes of a token with a known ID
(see "Token validation" below). The `opaqueId` is never reused — if a
paragraph is deleted and re-created, its opaque tokens get new IDs.

#### Sentinel character escaping

U+FFFC (Object Replacement Character) is the sentinel that distinguishes
tokens from text. It can appear in real content through:

- **Pasted text** from other applications
- **Imported .docx** content (rare but possible — Word uses U+FFFC for
  embedded OLE objects)
- **User typing** (effectively impossible on standard keyboards, but
  programmatic input could include it)

Rules:

1. **On import**: When reading run text from .docx, strip any literal
   U+FFFC characters from `w:t` text content. The ECMA-376 spec does
   not include U+FFFC in `w:t` — inline objects are child elements, not
   text characters. However, some producers (notably older OLE-embedding
   workflows) may emit U+FFFC as a placeholder in `w:t`. **This is a
   deliberate fidelity tradeoff**: we strip these because the sentinel
   must be unambiguous in the CRDT, and a literal U+FFFC in text would
   be indistinguishable from a token marker. In practice this affects
   near-zero real documents, but it is lossy for those edge cases.
   The stripped characters are logged (not silently discarded) so the
   import pipeline can report that fidelity was reduced.
2. **On paste** (`insertFromPaste` intent): The WASM `apply_intent`
   handler strips U+FFFC from the pasted text string before inserting
   into the CRDT. This runs before the CRDT transaction, so no
   sentinel can enter the text sequence from external content.
3. **On `insertText`**: Same strip — the intent handler normalizes all
   text input. This is a single `text.replace(/\uFFFC/g, '')` equivalent
   in Rust (`str.replace('\u{FFFC}', "")`).
4. **Serialization boundary**: When encoding CRDT state for persistence
   or sync, sentinels are part of the CRDT text sequence and travel with
   their token attributes. A sentinel without a token attribute is
   invalid state — `apply_update()` validation rejects it (see Token
   Validation above). A token attribute without a sentinel at that
   position is also invalid.
5. **DOM rendering**: Sentinel characters are never rendered as text in
   the DOM. The view model generator skips them (tokens become inline
   elements or are hidden). If a sentinel somehow appears in a DOM
   `textContent`, the position-mapping code treats it as zero-width.
6. **Copy from editor**: When the user copies a selection that spans
   tokens, the clipboard handler serializes tokens as their visual
   representation (e.g., a field shows its presentation text, an image
   becomes an `<img>` tag in HTML clipboard). Sentinels never appear in
   clipboard output. Pasting back into the editor creates new tokens
   only for recognized content (images via blob, etc.) — not by
   round-tripping sentinels.

No private-use-area characters are used. U+FFFC is the only sentinel.
This avoids the combinatorial complexity of multiple sentinel ranges
and keeps the escaping rules to one character.

#### Import (.docx → CRDT)

1. For each paragraph, walk runs sequentially
2. For each run:
   - Emit text characters into the CRDT RichText with formatting marks
   - For each non-text element (field, footnote ref, tab, break, image,
     unknown child): emit a sentinel character (U+FFFC) with the
     appropriate token attribute
3. Formatting marks span both text and sentinel characters (a bold
   footnote reference is bold)

#### Export (CRDT → .docx)

1. Read the character sequence with formatting marks and token attributes
2. Walk characters, tracking active formatting:
   - Regular characters: accumulate into current run's text
   - Sentinel characters: close current run (if any), expand token into
     OOXML element(s), start new run
3. When formatting changes: close current run, start new run
4. This is a pure function:
   `crdt_paragraph_to_elements(text, marks, tokens) -> Vec<RunOrElement>`

#### Scope constraint

Not all inline constructs are editable in v1. The initial supported set:

- **Editable**: text, formatting (bold/italic/etc), tabs, line breaks,
  inline images, paragraph properties
- **Preserved but not editable**: field codes, footnote/endnote refs,
  bookmark markers, comment markers, unknown children

"Preserved but not editable" means: the token exists in the CRDT, the
user's cursor can move past it (it occupies one character position),
but the editor UI does not provide controls to create or modify it.
The token survives editing and roundtrips correctly. Editing support
for individual token types is added incrementally.

#### Local intent enforcement (apply_intent rules)

Token immutability is not only a remote-validation concern — the local
`apply_intent()` path in WASM must also enforce it. Without this, a
user pressing Backspace next to a preserved token would delete it
through normal CRDT operations (which are local, not remote).

Rules enforced inside `apply_intent()` before the CRDT transaction:

1. **Cursor movement.** Arrow keys skip over preserved tokens as atomic
   units (one keypress = jump past the entire sentinel, not into it).
   The cursor never lands "inside" a token.
2. **Backspace / delete.** If the character immediately before the
   cursor (backspace) or after the cursor (delete) is a preserved
   sentinel, the deletion is **blocked** — the intent returns a no-op.
   The cursor does not move. This prevents accidental deletion of
   field codes, footnote refs, etc.
3. **Range deletion (selection delete, cut, replace-by-typing).** If
   the selected range contains preserved sentinels, two options:
   - **Strict (v1)**: Block the entire operation. The user must shrink
     their selection to exclude preserved tokens.
   - **Permissive (future)**: Delete only the text characters in the
     range, leaving preserved tokens in place. This is more complex
     (requires splitting the deletion around tokens) and deferred.
4. **Paste / insertText.** Text insertion adjacent to a preserved token
   is allowed — the new text goes before or after the sentinel, never
   replacing it. If a paste target range spans preserved tokens, rule 3
   applies.
5. **Formatting.** Applying formatting marks (bold, italic, etc.) to a
   range that includes preserved sentinels is allowed — formatting
   spans both text and tokens. This does not mutate the token's
   structural attributes.
6. **Editable tokens.** Tokens in the "editable" set (tabs, line breaks,
   inline images) can be deleted and inserted normally. Only tokens in
   the "preserved" set are protected by rules 1-4.
7. **Paragraph deletion.** Deleting an entire paragraph (e.g., selecting
   all content and pressing Backspace, or merging an empty paragraph
   into its neighbor) is always allowed, even when the paragraph
   contains preserved tokens. The protection rules above guard against
   _accidental_ individual token deletion — whole-paragraph removal is
   an intentional structural edit. The CRDT transaction removes the
   entire paragraph entry from the body array, and all its tokens
   (preserved or not) go with it.

These rules are enforced in the WASM `apply_intent()` handler, which
runs before any CRDT transaction is created. The TypeScript editor shell
does not need to implement these checks — it sends raw intents and the
WASM layer decides what to apply.

#### Token validation (collaboration trust boundary)

In collaborative mode, remote CRDT updates can contain token mutations.
Without validation, a malicious or buggy client could inject arbitrary
XML via opaque tokens. The WASM layer validates all token mutations
on `apply_update()`:

1. **Schema validation.** Every token attribute map must have a valid
   `type` field matching a known token type. Unknown types are rejected.
2. **Opaque token restrictions.** Opaque tokens (`type: "opaque"`) are
   the highest-risk because they carry raw XML. Rules:
   - **Immutable after import.** Each opaque token has an `opaqueId`
     (UUID v7, assigned on import). The WASM layer records the set of
     all `opaqueId` values present in the initial CRDT state. Any
     update that mutates the `xml` or `opaqueId` attributes of a
     token whose `opaqueId` is in this set is rejected (the entire
     update — see "Validation strategy" below). Deletion of the
     containing paragraph (which removes the sentinel character) is
     allowed.
   - **No new opaque tokens during collaborative editing.** Opaque
     tokens may only be created during `.docx` import — never during
     live editing sessions. This is enforced structurally, not by
     tracking which client "imported": the import pipeline runs as a
     batch operation that populates the initial CRDT state before any
     sync connection is established. Once the document enters
     collaborative mode (sync provider connected), the WASM
     `apply_update()` pre-validator rejects any update that creates
     new opaque tokens, regardless of which client sent it.

     Concrete flows:
     - **First open**: Client A imports .docx → CRDT state includes
       opaque tokens. This state is persisted as the initial snapshot.
       Client B joins, receives the snapshot (opaque tokens included).
       Neither client can create new opaque tokens from this point.
     - **Server-initiated import**: The server imports a .docx (e.g.,
       file upload endpoint) and stores the resulting CRDT snapshot.
       Clients receive it as initial state. Same rule: once any client
       connects, no new opaque tokens.
     - **Re-import (version replace)**: A new .docx replaces the
       document. This creates a new initial CRDT snapshot (with new
       opaque tokens). All connected clients receive a full state
       reset, not an incremental update. The old CRDT history is
       archived or discarded.
     - **Reconnect**: A client reconnecting to an existing session
       receives a state-vector diff or full snapshot — both contain
       opaque tokens from the original import. The reconnecting client
       cannot create new ones.

   - **Server-side enforcement.** If the server links the CRDT
     library, it runs the same pre-validation before relay (see
     "Validation strategy" below). A validating server is strongly
     recommended for production deployments.

3. **Field token integrity.** Field begin/code/separate/end tokens must
   maintain valid nesting (begin before code before separate before end,
   matching IDs). Malformed field sequences are rejected.
4. **Reference integrity.** Footnote/endnote ref tokens must reference
   IDs that exist in the CRDT footnotes/endnotes arrays. Dangling refs
   are rejected on `apply_update()`.

#### Validation strategy: pre-validate, then apply or reject whole update

CRDT wire formats (Yjs, Automerge, Loro) encode updates as causally
ordered operation sequences. Selectively dropping individual operations
corrupts state-vector consistency — the local replica believes it has
applied the update's clock entries, but the actual state diverges from
other replicas. Future merges produce incorrect results.

Therefore, validation is **all-or-nothing at the update level**:

1. **Pre-validate.** Before calling the CRDT library's native
   `apply_update()`, the WASM layer decodes the update into its
   operation list (library-specific: Yjs `decodeUpdate`, Automerge
   `getChanges`, Loro `decode_updates`) and checks each operation
   against the token rules above.
2. **If any operation fails validation**: reject the **entire update**.
   Do not apply it. Emit an `error` event to TypeScript with the
   rejection reason and the originating client ID (if available).
3. **If all operations pass**: apply the full update via the CRDT
   library's native `apply_update()`.
4. **Server-side (required for collaboration).** The collaboration
   server MUST link the CRDT library and run the same pre-validation
   before relaying updates. A rejected update is not broadcast to other
   clients. The server may disconnect the offending client.

   A dumb opaque relay cannot validate updates, and client-side-only
   rejection creates an unrecoverable state: the rejecting client falls
   behind the offending client's state vector, and the relay has no
   sanitized state to offer for recovery. **Therefore, dumb relays are
   supported only for single-user / development use** where there is no
   untrusted peer. Multi-client collaboration requires a validating
   server. The "Generic relay (prototyping only)" server option in the
   Server Implementation section is explicitly scoped to this.

This means a malicious client cannot cause silent data corruption or
replica divergence — the server blocks invalid updates before relay.
A malicious client can only get itself disconnected.

### What Stays Outside the CRDT

Some OOXML data is not editable in the CRDT and is preserved via the
**original immutable Document** (see Export Strategy below):

- **RawXmlNode — paragraph-level and above** (unknown children of
  `w:body`, `w:p`, `w:tbl`, `w:tc`, etc.): Live on the original
  Document's paragraph/table objects. Never detached, never stored
  as a blob. On export, the clone-and-patch strategy preserves them
  in place — we replace runs inside dirty paragraphs but never touch
  the paragraph element's own unknown children.
- **RawXmlNode — run-level** (unknown children inside `w:r`): These
  do NOT stay on the original Document because dirty paragraphs get
  their runs replaced on export. Instead, run-level unknowns are
  captured on import as `opaque` inline tokens in the CRDT RichText
  (see "Unknown Elements" under Key Mapping Challenges). They survive
  editing via the token model, not via the clone.
- **Style definitions**: Loaded into a read-only `StyleRegistry` on
  import. Style application (setting a paragraph's `style_id`) is a
  CRDT operation; style definition editing is not in scope.
- **Theme data, document properties, custom XML parts**: Preserved
  on the original Document, passthrough on export via clone.

### Image Storage: Out-of-Band Blobs

Image binaries are NOT stored in the CRDT. A 5MB pasted image inside
the CRDT becomes a 5MB replicated update in every client's history —
this kills sync performance and bloats storage.

Instead, images use a side-channel blob store:

```
CRDT images map:
  "sha256:abc123..." → { contentType: "image/png", width: 914400, height: 612000 }

Blob store (outside CRDT):
  Client-side:  "sha256:abc123..." → <actual binary bytes>
  Server-side:  ("doc_42", "sha256:abc123...") → <actual binary bytes>
```

The CRDT stores only metadata (content hash, dimensions, content type).
The actual bytes live in a blob store:

- **Client-side**: IndexedDB blob store, keyed by content hash
- **Server-side**: S3/R2/blob storage, keyed by `(document_id, content_hash)`
- **Original document**: Images from the imported .docx are retained on
  the immutable Document and available via `image_data_uri(index)`

When a client references an image hash it doesn't have locally, it
fetches the blob from the server's blob store API (not peer-to-peer).
Server-mediated fetch is the only supported model — it simplifies auth
(one token check per request), works when the uploading client is
offline, and avoids WebRTC complexity. Content-hash keying gives
automatic deduplication and integrity verification.

On export, the blob store is consulted to populate the .docx image parts.

#### Blob store lifecycle and security

**Document scoping.** Blobs are scoped to a document ID, not globally
shared. The storage key is `(document_id, content_hash)`, not just
`content_hash`. This prevents:

- Cross-tenant blob leakage (one user's images visible to another)
- Hash collision attacks (unlikely with SHA-256 but defense in depth)

**Authorization.** Blob fetch requests require the same auth token as
the document's WebSocket connection. The server validates that the
requesting client has read access to the document before serving a blob.

**Orphan garbage collection.** Blobs can become orphaned when:

- An image is pasted then immediately deleted (undo)
- A paragraph containing an image is deleted
- CRDT compaction removes the image reference from history

GC strategy:

1. After each CRDT compaction (or periodically), scan all image refs
   in the current CRDT state.
2. Compare against blobs stored for this document.
3. Blobs not referenced in the current state AND not referenced in
   any retained CRDT snapshot are eligible for deletion.
4. Grace period: don't delete orphans immediately — wait T hours
   (e.g., 24h) in case of undo, late-arriving updates, or clients
   with stale state.
5. Client-side (IndexedDB): same logic, shorter grace period (1h).

**Size limits.** Per-document blob budget (e.g., 100MB). Individual
blob size limit (e.g., 10MB per image). Enforced on upload — the
server rejects blobs exceeding the limit.

---

## CRDT Library Candidates

### yrs (Y-CRDT for Rust)

- **Crate**: `yrs`
- **What**: Official Rust port of Y.js. Sequence CRDT with rich text
  attributes, arrays, maps.
- **Rich text**: `YText` with inline formatting attributes. Maps well
  to docx runs.
- **Positions**: `StickyIndex` for stable cursor positions.
- **Awareness**: Built-in via `y-sync` crate.
- **Undo**: Built-in `UndoManager`, scoped per client.
- **WASM**: First-class support via `y-wasm`.
- **Sync**: `y-sync` crate implements the Yjs sync protocol. State
  vectors + `encode_diff_v2` / `encode_state_as_update_v2` for
  incremental binary diffs.
- **Ecosystem**: Wire-compatible with Y.js, y-swift, y-py. Existing
  providers: `y-websocket`, `y-indexeddb` (JS-side). Hocuspocus server.
- **Production users**: AppFlowy, BlockSuite, AFFiNE, Huly.
- **Tradeoffs**: Largest ecosystem but Y-CRDT table merge semantics
  may not map 1:1 to OOXML. GC tuning needed for large docs.

### automerge

- **Crate**: `automerge`
- **What**: JSON-like CRDT. Nested maps, lists, and text.
- **Rich text**: `Text` type with `mark()` for formatting spans (2.x+).
- **Positions**: `Cursor` type for stable sequence positions.
- **Awareness**: Not built-in; needs custom implementation.
- **Undo**: Not built-in; implement via `PatchLog` inverse.
- **WASM**: `@automerge/automerge` npm package, `automerge-wasm` crate.
- **Sync**: Built-in sync protocol with `SyncState` + `generate_sync_message`
  / `receive_sync_message`. `PatchLog` for incremental change observation.
- **Ecosystem**: `automerge-repo` for networking/storage. Smaller than
  Y.js but well-maintained.
- **Production users**: Ink & Switch projects, some commercial apps.
- **Tradeoffs**: Different merge semantics (list CRDT, no "position
  relative to deleted items"). Potentially simpler mental model. Smaller
  ecosystem means more DIY for awareness/undo.

### loro

- **Crate**: `loro`
- **What**: High-performance CRDT framework. Rich type system: movable
  tree, list, map, rich text.
- **Rich text**: `LoroText` with `mark()` API. Supports expand behavior
  (before/after/both/none) for marks — useful for formatting boundaries.
- **Positions**: `Cursor` with `Side::Left` / `Side::Right`.
- **Awareness**: Built-in awareness/presence support.
- **Undo**: Built-in undo manager.
- **WASM**: First-class support via `loro-wasm`.
- **Sync**: `export(ExportMode::updates_in_range(...))` for incremental
  diffs. State vectors for catch-up.
- **Ecosystem**: Custom protocol (no Y.js interop). Own sync/storage
  primitives.
- **Production users**: Newer, actively developed. Growing adoption.
- **Tradeoffs**: Movable tree type could be useful for document structure
  (reordering paragraphs, drag-and-drop). Newest of the mature options.
  Custom protocol means building providers from scratch.

### diamond-types

- **Crate**: `diamond-types`
- **What**: High-performance text CRDT by Joseph Gentle. Run-length
  encoded operations.
- **Rich text**: None. Plain text only.
- **Positions**: None built-in.
- **Awareness/Undo/WASM**: None/none/compiles but no bindings.
- **Tradeoffs**: Fastest raw text CRDT by benchmarks. Would need
  everything beyond plain text built from scratch. Better as a
  building block or benchmark reference than a complete solution.

### cola

- **Crate**: `cola_crdt`
- **What**: Lightweight text CRDT by Nomic AI.
- **Rich text**: None. Plain text only.
- **Positions/Awareness/Undo**: None.
- **Tradeoffs**: Smallest footprint. Same limitation as diamond-types:
  no rich text, no formatting, no structured types.

---

## View Model: CRDT-Driven Updates

### On Initial Load

Full conversion: CRDT state → `DocViewModel` (reuse existing
`convert_document` logic, adapted to read from CRDT instead of
`offidized_docx::Document`).

### On Edit

Subscribe to CRDT change events. Each CRDT library emits typed
notifications about what changed:

- **yrs**: `TransactionMut` events, `ObserveDeep` on shared types
- **automerge**: `PatchLog` with `SpliceText`, `PutMap`, `Insert`, etc.
- **loro**: `Subscription` with `DiffEvent` per container

Map change events to `ViewPatch` values:

```rust
pub enum ViewPatch {
    ReplaceParagraph { index: usize, model: ParagraphModel },
    InsertParagraph { index: usize, model: ParagraphModel },
    RemoveParagraph { index: usize },
    ReplaceTable { index: usize, model: TableModel },
    InsertTable { index: usize, model: TableModel },
    RemoveTable { index: usize },
    UpdateSelection { selections: Vec<RemoteCursor> },
}
```

TypeScript receives patches (as JSON) and applies them to the DOM.
No full re-render per keystroke.

### Diffing

Use CRDT-native diffs, not generic text diff libraries:

- **yrs**: `encode_state_vector` + `encode_diff_v2`. Update events
  on shared types give you exactly which items changed.
- **automerge**: `PatchLog` + sync protocol for incremental patches.
  `diff()` between document states.
- **loro**: `export(ExportMode::updates_in_range(...))`. Subscription
  diffs give typed change events per container.

The CRDT already knows exactly what changed — don't re-derive it.

---

## Persistence: TypeScript-Side Providers

Storage and networking live entirely in TypeScript. The WASM boundary
is a thin FFI that passes binary blobs:

### WASM FFI Surface

```rust
// Exposed to JS via wasm_bindgen:

/// Get the current state vector (for sync catch-up).
pub fn encode_state_vector(&self) -> Vec<u8>;

/// Encode all state as a binary update (for full save).
pub fn encode_full_update(&self) -> Vec<u8>;

/// Encode only changes since a given state vector (for incremental sync).
pub fn encode_diff(&self, remote_state_vector: &[u8]) -> Vec<u8>;

/// Apply a remote update (from another client or from storage).
pub fn apply_update(&mut self, update: &[u8]) -> Result<(), JsValue>;

/// Export current CRDT state as .docx bytes.
pub fn snapshot_to_docx(&self) -> Result<Vec<u8>, JsValue>;

/// Awareness: encode local cursor/user state.
pub fn encode_awareness(&self) -> Vec<u8>;

/// Awareness: apply remote cursor/user state.
pub fn apply_awareness(&mut self, update: &[u8]);
```

### TypeScript Providers

All use browser APIs. Each implements a simple interface:

```typescript
interface DocProvider {
  /** Called when WASM produces a CRDT update. */
  onLocalUpdate(update: Uint8Array): void;
  /** Called to deliver a remote update to WASM. */
  onRemoteUpdate: ((update: Uint8Array) => void) | null;
  /** Connect/start the provider. */
  connect(): Promise<void>;
  /** Disconnect/stop. */
  disconnect(): void;
}
```

#### IndexedDB Provider

- Offline-first. Stores CRDT updates incrementally + periodic snapshots.
- On load: replay updates on top of latest snapshot.
- Use raw IndexedDB API or `idb` (tiny wrapper).
- Compaction policy (see below).
- Reference implementations: `y-indexeddb` (Yjs), `automerge-repo`
  IndexedDB adapter (Automerge).

#### File System Access Provider

- Uses File System Access API (`showSaveFilePicker`, `FileSystemFileHandle`).
- Calls `snapshot_to_docx()` to save real `.docx` files to disk.
- Fallback: download via `<a>` blob URL if FS Access unavailable.
- Not incremental — full save on each persist.

#### HTTP Provider

- POST/GET `.docx` bytes to a server endpoint.
- Calls `snapshot_to_docx()` for the upload payload.
- No realtime — manual or auto-save on interval.

#### WebSocket Sync Provider

- Binary WebSocket frames for CRDT updates.
- On connect: exchange state vectors, send diff, apply remote diff.
- Ongoing: relay incremental updates bidirectionally.
- Awareness messages for cursor positions (if `awareness` feature enabled).
- Reconnection with state vector catch-up (stateless reconnect).
- Libraries: raw `WebSocket` API, or `partysocket` (~2kb, adds
  reconnection + room multiplexing).
- Reference implementations: `y-websocket` (Yjs), `automerge-repo`
  WebSocket adapter (Automerge).

### Update-Log Compaction Policy

CRDT update logs grow unboundedly without compaction. Both IndexedDB
and server-side storage need explicit retention strategies:

#### Client-side (IndexedDB Provider)

- **Snapshot interval**: After every N updates (e.g., 500) or M bytes
  of accumulated updates (e.g., 1MB), write a full CRDT state snapshot.
- **Retention**: Keep only the latest snapshot + updates since that
  snapshot. Delete older snapshots and their associated update runs.
- **On load**: Load latest snapshot, replay subsequent updates.
- **Trigger**: Compaction runs on `requestIdleCallback` or after a
  save, never on the hot edit path.

#### Server-side

- **Snapshot interval**: After N updates per document (e.g., 1000) or
  when a document has been idle for T seconds (e.g., 60), compact all
  stored updates into a single snapshot.
- **Retention**: Keep the latest snapshot + updates since snapshot +
  optionally one prior snapshot (for rollback). Delete everything older.
- **Acknowledgment**: Before compacting, ensure all connected clients
  have received updates up to the compaction point (their state vectors
  are >= the snapshot's clock). Disconnected clients that reconnect
  with a state vector older than the oldest retained snapshot must do
  a full state transfer instead of incremental catch-up.
- **Size budget**: Alert/warn if a document's total stored state
  exceeds a threshold (e.g., 50MB). This catches pathological cases
  like rapid image-heavy editing.

#### CRDT-library-specific compaction

- **yrs**: `Doc::gc()` for garbage-collecting tombstones.
  `encode_state_as_update_v2` for full snapshot encoding.
- **automerge**: `Automerge::save()` produces a compact single-blob
  snapshot. `compact()` for in-memory compaction.
- **loro**: `export(ExportMode::Snapshot)` for full snapshot.
  `export(ExportMode::updates_in_range(...))` for delta since a version.

---

## Collaboration Server

Not optional — the protocol, state-vector catch-up, and auth semantics
need to be designed from day 1 even if the server ships later. The client
sync protocol is the same whether talking to a production server or a
local dev relay.

### Responsibilities

1. **Room management.** Each document = one room. Clients join/leave.
2. **Update relay.** Broadcast CRDT updates to all room participants.
3. **Awareness relay.** Broadcast cursor/selection/user-info.
4. **Persistence.** Store CRDT updates for late joiners. Periodic
   compaction into snapshots.
5. **Auth.** Validate identity and document access on connect.
6. **.docx export.** Server-side snapshot via offidized-docx (optional
   but valuable — lets you generate download links without a client).

### Server Protocol

The CRDT library choice determines the wire protocol. Different CRDT
stacks (Yjs, Automerge, Loro) have incompatible sync protocols — they
do not share state vector formats, update encoding, or message framing.
Attempting to abstract across them adds complexity with no benefit.

**One CRDT library is chosen at build time.** The `editing` feature flag
compiles in exactly one library. The sync protocol, server, and providers
are built for that library's wire format. If the library is later swapped,
the server and providers are also swapped — this is an infrequent,
deliberate decision, not a runtime concern.

The protocol shape (regardless of library) follows this pattern:

1. **Auth**: Token exchange on connect (before sync starts).
2. **SyncStep1**: Client sends its state vector (library-specific encoding).
3. **SyncStep2**: Server responds with diff (library-specific encoding).
4. **Update**: Either direction — incremental update bytes (library-specific).
5. **Awareness**: Cursor/presence updates (library-specific or custom).

For a dumb relay server that doesn't understand CRDT internals, messages
1-4 are opaque binary frames tagged by type byte. For a smart server
(compaction, server-side export), it links the same CRDT library.

The CRDT library candidates section above exists as a decision record
for choosing the library. Once chosen, the remaining sections assume
that library's protocol.

### Server Implementation Options

#### Rust-native: axum + CRDT library

- `axum` for HTTP/WebSocket.
- Same CRDT library server-side for compaction, validation, export.
- Can run offidized-docx for server-side `.docx` snapshot generation.
- Full control over protocol, persistence, and auth.
- If using yrs: `y-sync` implements the Yjs sync protocol directly.
- Tradeoff: most work to build, most capability.

#### Hocuspocus (Yjs ecosystem only)

- JavaScript/TypeScript WebSocket server by TipTap.
- Mature: auth hooks, webhooks, multiple storage backends.
- Open source (MIT). Only works with Yjs-compatible clients.
- Runs on Node, Bun, Deno.
- Tradeoff: JS runtime (can't run offidized-docx server-side),
  locks you into Yjs ecosystem.

#### y-websocket (Yjs ecosystem only)

- Minimal reference server for Yjs sync protocol.
- ~200 lines. Good for prototyping.
- Tradeoff: no auth, no persistence hooks, Yjs-only.

#### Cloudflare Durable Objects

- Each document = one Durable Object.
- WebSocket Hibernation API for cost efficiency.
- KV/R2 for CRDT update + blob persistence.
- Pairs with a generic Cloudflare Worker or relay deployment pattern.
- Limits: 128MB memory per DO, 1000 WebSocket connections per DO.
- Tradeoff: vendor lock-in, JS runtime, but excellent scaling model.

#### PartyKit / Partyserver

- Managed WebSocket rooms. Built for multiplayer state sync.
- Relays opaque CRDT bytes without interpreting them.
- Tradeoff: managed service dependency, can't compact server-side.

#### Generic relay (single-user dev / prototyping only)

- Server doesn't understand CRDT — just relays binary frames per room.
- Any WebSocket server works (20-line Bun script).
- Can't compact, validate, or export. Growth is unbounded.
- **Not suitable for multi-client collaboration.** Without server-side
  validation, a malicious or buggy client can inject invalid updates
  that cause other clients to reject and fall out of sync with no
  recovery path (see "Validation strategy" above).
- Use only for single-user sync (e.g., syncing between a user's own
  browser tabs) or local development.

---

## Threading Model: Web Worker Boundary

Several operations are too expensive for the main thread:

- **`snapshot_to_docx()`**: Clone-and-patch export. Deep-clones the
  original Document, patches dirty paragraphs, serializes to ZIP.
  Can take 100ms+ on large documents.
- **CRDT compaction**: GC, snapshot encoding, state merge. CPU-heavy.
- **Import**: Parsing a large .docx (ZIP decompression, XML parsing,
  CRDT population). Can take seconds.
- **Full view model generation**: Initial `convert_document` on a
  1000-paragraph document.

Running these on the UI thread will cause input lag and render jank.

### Architecture

```
Main Thread                          Worker Thread
┌──────────────────────┐             ┌──────────────────────┐
│ Editor Shell (TS)    │             │ WASM Module          │
│ contenteditable      │  postMessage│ CRDT Doc             │
│ Selection, IME       │◄───────────►│ View model gen       │
│ DOM patching         │  (binary)   │ Import / Export      │
│ Providers (WS, IDB)  │             │ Compaction           │
└──────────────────────┘             └──────────────────────┘
```

The WASM module runs in a **dedicated Web Worker**. The main thread
communicates via `postMessage` with `Transferable` (zero-copy) binary
buffers. This keeps all CPU-heavy work off the main thread.

### Message protocol (main ↔ worker)

```typescript
// Main → Worker
type WorkerRequest =
  | { type: "load"; data: ArrayBuffer } // import .docx
  | { type: "intent"; intent: EditIntent } // user edit
  | { type: "applyUpdate"; update: ArrayBuffer } // remote CRDT update
  | { type: "applyAwareness"; data: ArrayBuffer } // remote cursors
  | { type: "save" } // trigger .docx export
  | { type: "encodeUpdate" } // get local CRDT update
  | { type: "encodeStateVector" } // for sync handshake
  | { type: "compact" }; // trigger compaction

// Worker → Main
type WorkerResponse =
  | { type: "viewPatches"; patches: ViewPatch[] } // DOM updates
  | { type: "fullModel"; model: DocViewModel } // initial render
  | { type: "localUpdate"; update: ArrayBuffer } // send to provider
  | { type: "awareness"; data: ArrayBuffer } // send to provider
  | { type: "saved"; docx: ArrayBuffer } // .docx bytes
  | { type: "stateVector"; sv: ArrayBuffer }
  | { type: "error"; message: string };
```

### What stays on main thread

- DOM manipulation (contenteditable, selection, cursor rendering)
- Browser event handling (keyboard, mouse, IME, clipboard)
- Persistence providers (WebSocket, IndexedDB) — they need main-thread
  or worker access; WebSocket works from either, IndexedDB from either
- `postMessage` dispatch

### What runs in the worker

- All WASM execution (CRDT mutations, view model generation, export)
- CRDT compaction and GC
- Import (parse .docx → CRDT)

### Latency budget

For interactive editing (keystroke → DOM update):

1. Main thread: capture `beforeinput`, encode intent → `postMessage` (~1ms)
2. Worker: apply CRDT transaction, generate ViewPatch → `postMessage` (~1-5ms)
3. Main thread: apply ViewPatch to DOM (~1ms)

Total: ~3-7ms. Well within the 16ms frame budget. The `postMessage`
overhead is ~0.1ms for small messages with `Transferable`.

For export (`snapshot_to_docx`): runs entirely in the worker. Main
thread is never blocked. The `saved` response carries the `.docx`
`ArrayBuffer` as a `Transferable` (zero-copy transfer).

### Fallback: no Worker

For environments where Workers aren't available (some embedded
contexts), the WASM module can run on the main thread. The same
message protocol is used but with synchronous function calls instead
of `postMessage`. Export and compaction should be deferred to
`requestIdleCallback` in this mode.

---

## TypeScript Editor Shell

### Responsibility Split

| Concern                             | Where                    |
| ----------------------------------- | ------------------------ |
| `contenteditable` div               | TypeScript               |
| Selection API, cursor tracking      | TypeScript               |
| IME composition events              | TypeScript               |
| Clipboard (copy/paste)              | TypeScript               |
| Keyboard shortcuts (Ctrl+B, Ctrl+Z) | TypeScript               |
| Mouse events → cursor positioning   | TypeScript               |
| Toolbar state queries               | TypeScript → WASM        |
| Remote cursor rendering             | TypeScript               |
| DOM patching from ViewPatch         | TypeScript               |
| CRDT state (all mutations)          | WASM                     |
| View model generation               | WASM                     |
| Undo/redo                           | WASM (CRDT undo manager) |
| .docx import/export                 | WASM                     |
| Position encoding/decoding          | WASM                     |
| Awareness encoding                  | WASM                     |
| Persistence / networking            | TypeScript               |

### contenteditable Strategy Options

#### Option A: Controlled contenteditable

Intercept `beforeinput` events. Prevent default on everything except
IME composition. Map `InputEvent.inputType` to editing intents. After
WASM processes the intent, reconcile DOM from view patches.

- **Pro**: Native text input, cursor, IME, accessibility.
- **Con**: Browser quirks. Each browser handles `beforeinput` differently.
  Reconciliation between WASM-driven DOM updates and browser's native
  mutations is the hardest part.
- **Used by**: ProseMirror, TipTap, Lexical, Slate.

#### Option B: Hidden textarea + custom rendering

Render document as non-editable DOM. Hidden `<textarea>` captures input.
Custom cursor (CSS animated caret), custom selection highlighting.

- **Pro**: Full control. No browser contenteditable quirks.
- **Con**: Must implement cursor rendering, selection painting, IME
  overlay positioning, scrolling, text measurement.
- **Used by**: CodeMirror 6, Monaco, Google Docs (canvas variant).

#### Option C: Canvas rendering

Entire document on `<canvas>`. Hidden textarea for input. Full custom
everything.

- **Pro**: Maximum rendering control. Same pattern as xlview.
- **Con**: Must implement text layout, line breaking, paragraph flow.
  Accessibility requires parallel ARIA live region. Not practical for
  flowing rich text with variable-width content.
- **Used by**: Google Docs (newer), Figma (for text in shapes).

**Recommendation**: Option A for docx. Rich text with flowing paragraphs,
tables, and images benefits from the browser's layout engine. The
contenteditable complexity is real but bounded, and every production
rich text editor has solved it. Option B is viable if contenteditable
proves too painful — CodeMirror 6 demonstrates it works well.

### Intent Passing

TypeScript captures browser events, resolves selection to CRDT-relative
positions (via `encode_position`), and passes structured intents to WASM:

```typescript
interface EditIntent {
  type:
    | "insertText"
    | "deleteContentBackward"
    | "deleteContentForward"
    | "insertParagraph"
    | "insertLineBreak"
    | "formatBold"
    | "formatItalic"
    | "formatUnderline"
    | "insertFromPaste"
    | "deleteByCut"
    | "historyUndo"
    | "historyRedo";
  data?: string;
  /** Opaque CRDT-relative position bytes, not flat offsets. */
  anchor: Uint8Array;
  focus: Uint8Array;
}
```

WASM processes the intent as a CRDT transaction, emits view patches,
returns them to TypeScript for DOM application.

### DOM ↔ CRDT Position Mapping

Each rendered paragraph has a `data-para-id` attribute (the app-level
UUID). On selection change, TypeScript:

1. Reads the browser `Selection` (anchor node, anchor offset, focus node,
   focus offset). The DOM `Selection` API returns offsets in UTF-16 code
   units for text nodes — this is the canonical unit.
2. Walks up to find the containing `[data-para-id]` element.
3. Sums `textContent.length` (which is UTF-16 code units) of preceding
   text nodes within the paragraph to compute the paragraph-level offset.
4. Calls `encode_position(paraId, utf16Offset)` → opaque `Uint8Array`.

The UTF-16 code unit is the canonical unit at the TypeScript ↔ WASM
boundary. WASM converts to the CRDT's internal indexing inside
`encode_position`. TypeScript never computes offsets in any other unit.

This opaque position is what gets stored, passed to WASM, and synced
via awareness. It survives concurrent edits because it's CRDT-relative.

---

## Import / Export Pipeline

### Import (.docx → CRDT)

```
.docx bytes
    │
    ├──► offidized_docx::Document::from_bytes()
    │        │
    │        ├──► Retained as immutable snapshot (for export)
    │        │
    │        └──► Walk document structure → populate CRDT:
    │              - For each paragraph: create CRDT Array entry (Map)
    │                - Assign UUID, flatten runs → RichText with marks
    │                - Set paragraph properties
    │                - Record mapping: UUID → original paragraph index
    │              - For each table: create CRDT Array entry
    │              - For images: hash bytes → register in CRDT images map
    │                  (blobs go to side-channel store, not CRDT)
    │
    └──► CRDT Doc initialized + immutable Document retained
```

The original `Document` object is kept alive (immutable) alongside the
CRDT. It owns paragraph-level and above `RawXmlNode` unknown elements,
style definitions, themes, document properties, and every other OOXML
part that the CRDT doesn't model. These stay in their original parse
tree positions — nothing is extracted into blobs or keyed by path.

The one exception: **run-level** unknown children (`RawXmlNode` inside
`w:r`) are extracted into the CRDT as `opaque` inline tokens during
import, because dirty paragraphs get their runs fully replaced on
export (see "Unknown Elements" under Key Mapping Challenges).

### Export (CRDT → .docx): Clone-and-Patch

This is the critical strategy for roundtrip fidelity. The export does
NOT build a new `Document` from scratch. It clones the original and
patches only what the CRDT says changed:

```
Immutable original Document
    │
    ├──► Clone (deep copy)
    │
    ▼
Cloned Document
    │
    ├──► For each CRDT body item:
    │      - If UUID matches an original paragraph AND paragraph is dirty:
    │          → Replace runs on the cloned paragraph
    │            (read CRDT RichText + inline tokens → expand to Runs
    │             and inline elements; opaque tokens re-emit raw XML)
    │          → Update paragraph properties from CRDT
    │          → RawXmlNode children of the PARAGRAPH element (not runs)
    │            on the clone are UNTOUCHED
    │          → Run-level unknown children survive via opaque tokens
    │      - If UUID matches an original paragraph AND NOT dirty:
    │          → Skip entirely (clone already has correct content)
    │      - If UUID has no original match (new paragraph):
    │          → Insert new Paragraph at correct position in clone
    │      - If original paragraph has no CRDT match (deleted):
    │          → Remove from clone
    │
    ├──► Resolve image references:
    │      - New images: fetch from blob store, add to clone's images
    │      - Existing images: already present on clone
    │
    ├──► Styles, themes, doc properties, custom XML:
    │      Already on the clone, untouched
    │
    ▼
Cloned Document.save() → .docx bytes
```

Why this works:

- **Paragraph-level and above unknown elements survive** because they
  live on the original `Paragraph` / `Table` objects and are carried
  through the clone. We only replace the runs inside dirty paragraphs —
  the paragraph element's own unknown children are preserved.
- **Run-level unknown elements survive** via the opaque token path:
  they were extracted into CRDT `opaque` tokens on import (with stable
  `opaqueId`), and on export the token's `xml` attribute is re-emitted
  into the newly generated runs. The original `Run` objects on the clone
  are replaced, but the unknown content they carried is reconstructed
  from the CRDT.
- **Structural edits are safe.** A new paragraph inserted between two
  existing ones creates a new `Paragraph` with no unknown children
  (correct — it never had any). A deleted paragraph removes it from
  the clone. Unchanged paragraphs keep their full original state.
- **Relationship IDs, content types, and ZIP structure** are handled
  by `offidized_docx::Document::save()` which already reconstructs
  these from the document tree.

This is the same pattern as xlview's `zip_patcher`: copy unmodified
parts verbatim, re-serialize only dirty parts. The difference is that
xlview operates at the ZIP entry level, while docview operates at the
paragraph level within the document.xml part.

### Dirty Tracking

The CRDT maintains a set of dirty paragraph UUIDs (paragraphs whose
content or properties have changed since the last export). On export,
only these paragraphs are patched on the clone. After a successful
export, the dirty set is cleared.

For new documents (created from scratch, not imported), there is no
original Document to clone — the export builds a fresh Document from
CRDT state. This is fine because there are no unknown elements to
preserve.

---

## Key Mapping Challenges

### Rich text ↔ Runs

See "Rich Text ↔ Runs Mapping" section above for the full inline
token model. The short version:

- CRDT rich text contains text characters + U+FFFC sentinel characters
- Sentinels carry structured token attributes (field begin/code/separate/end,
  footnote refs, tabs, breaks, images, opaque XML)
- Complex fields use multi-token representation (begin/code/separate/end
  with shared field ID) — no lossy collapsing to single tokens
- Formatting marks span both text and sentinels
- Import flattens runs to text+tokens; export coalesces back to runs
- The `opaque` token type is the escape hatch for unknown run children
- U+FFFC is stripped from all user input (paste, typing) — see
  "Sentinel character escaping" for full rules

### Tables

Concurrent structural changes (add/remove rows/columns with merged cells)
are hard in any CRDT. Pragmatic approach:

- Each cell's text content is an independent CRDT RichText (concurrent
  cell edits merge cleanly).
- Structural changes (add row, merge cells) are coarse-grained: treated
  as a single CRDT transaction. If two users add a row simultaneously,
  both rows appear (CRDT array insert semantics).
- Merged cells: store merge info as cell properties. Concurrent merge
  and split may conflict — last-write-wins at the cell property level.

### Images

Image binaries live in a side-channel blob store, not in the CRDT
(see "Image Storage: Out-of-Band Blobs" above). The CRDT stores only
references:

```
CRDT images map: { "sha256:abc123..." → { contentType, width, height } }
Inline image in rich text: sentinel char with attribute
  { imageRef: "sha256:abc123...", width: 914400, height: 612000 }
```

Content-hash keying gives automatic deduplication. When a new image is
pasted, the client hashes the bytes, stores them in the local blob store,
and inserts only the hash + metadata into the CRDT. Other clients see
the hash in the CRDT update, fetch the blob from the server's blob API,
and cache it locally.

### Unknown Elements (RawXmlNode)

Unknown elements exist at two levels, handled differently:

**Paragraph-level and above** (unknown children of `w:body`, `w:p`,
`w:tbl`, `w:tc`, etc.): These live on the original immutable Document.
They survive export via clone-and-patch — we only replace runs inside
dirty paragraphs, never the paragraph element's own unknown children.

**Run-level** (unknown children inside `w:r`): These cannot survive on
the original Document because we replace all runs on dirty paragraphs.
Instead, they are captured on import as `opaque` inline tokens (sentinel
characters with `{ type: "opaque", opaqueId: "<uuid>", xml: "<w:...>...</w:...>" }` attributes
in the CRDT RichText). On export, opaque tokens are re-emitted as raw
XML inside the generated run. This means run-level unknown elements
survive editing and roundtrip correctly, even when the containing
paragraph is modified.

### Comments and Tracked Changes

Two approaches (open question):

1. **In the CRDT**: Comments as a separate CRDT map, keyed by ID, with
   anchor positions (CRDT-relative). Tracked changes as marks on the
   rich text (insert/delete tracking).
2. **Outside the CRDT**: Preserved on the original immutable Document
   (same as styles/themes), carried through clone-and-patch on export.
   Not editable during collaborative sessions. Simpler but means
   comments can't be added/resolved collaboratively.

### Style Inheritance

Paragraph and run formatting can inherit from styles. The editor needs
to resolve effective formatting for toolbar state (is this text bold?).
The CRDT stores explicit formatting overrides; style resolution is a
read-only computation at view-model generation time using the imported
`StyleRegistry`.

---

## Phased Implementation

### Phase 1: Editing with CRDT (single client)

1. Choose CRDT library. Add behind `editing` feature flag.
2. Build import pipeline: `.docx` → CRDT doc (via offidized-docx parse).
3. Build view model from CRDT state (adapt existing `convert_document`).
4. Build `DocEdit` WASM struct: wraps CRDT doc + view model.
5. Implement CRDT-relative position encoding/decoding.
6. Build TypeScript editor shell: contenteditable + intent passing.
7. Wire: keystroke → intent → CRDT transaction → view patch → DOM.
8. Undo/redo via CRDT undo manager (free from the library).
9. Build export pipeline: CRDT → offidized-docx → .docx bytes.
10. Save/load: full `.docx` round-trip.

### Phase 2: Incremental updates + providers

1. Subscribe to CRDT change events → emit ViewPatch (not full re-render).
2. Build IndexedDB provider (offline persistence of CRDT state).
3. Build File System Access provider (save as .docx to disk).
4. Build WebSocket sync provider (client side only — talks to any
   server that speaks the chosen CRDT library's wire protocol).
5. Awareness protocol: encode/apply local and remote cursor state.
6. Render remote cursors in TypeScript (colored carets + name labels).

### Phase 3: Server + production hardening

1. Deploy validating collaboration server (Rust/axum with linked CRDT
   library, ecosystem-specific server, or Cloudflare DO). Multi-client
   collaboration requires server-side update validation — a generic
   relay is not sufficient (see "Validation strategy" above).
2. Room management, auth, persistence.
3. Server-side .docx export (snapshot without a client).
4. Large document performance: virtualized rendering (only render
   visible pages), lazy CRDT loading for large docs.
5. Table editing, image handling, copy/paste fidelity.

---

## Confirmed Assumptions

These decisions are locked:

1. **One CRDT library per build.** The library choice determines the
   wire protocol, server, and providers. No runtime CRDT abstraction.
2. **Inline token model for non-text content.** Footnote refs, bookmark
   markers, images, and unknown run children are single-sentinel tokens.
   Complex field codes use multi-token representation (begin/code/
   separate/end with shared ID) preserving original run structure.
   Opaque tokens are the escape hatch. U+FFFC is the only sentinel;
   stripped from all user input at the intent handler boundary.
3. **Editable inline scope for v1.** Text, formatting, tabs, line
   breaks, inline images, paragraph properties. Everything else is
   preserved-but-not-editable via tokens.
4. **Perf and fidelity gates required before GA.** Large real-world
   docs (.docx files from production Word/Google Docs) must be tested
   for roundtrip fidelity and edit performance before release.

## Open Questions

- **CRDT library choice.** Needs prototyping. Key differentiators: rich
  text mark semantics (especially: how do marks interact with sentinel
  characters?), stable position API, WASM binary size, UTF-16 ↔
  internal offset conversion cost, ecosystem providers. Build a small
  spike with the top candidates (yrs, loro, automerge).
- **Table structural edits.** Concurrent add-row + merge-cells is
  unsolved in most CRDTs. May need custom conflict resolution or
  coarse-grained locking at the table level.
- **Comments in CRDT.** Worth the complexity? Or keep comments as a
  separate channel outside the document CRDT? Collaborative comment
  resolution (mark as resolved) is a strong user expectation.
- **Large document performance.** Thousands of paragraphs need
  virtualized rendering. The paginator exists but needs to work with
  incremental view patches and the Worker message protocol.
- **contenteditable vs hidden textarea.** Prototype both. The
  contenteditable approach is standard but notoriously painful. A hidden
  textarea gives more control at the cost of building cursor/selection
  rendering.
- **Clone-and-patch for large documents.** Deep-cloning a large Document
  on every export may be expensive. May need copy-on-write or structural
  sharing. Profile with real-world documents before optimizing.
- **Field code evaluation.** Complex fields use the multi-token model
  (begin/code/separate/end) which preserves structure. Open question:
  should the editor re-evaluate field results (e.g., update PAGE numbers
  after reflow) or always show the imported presentation text? Evaluation
  is hard (some fields reference external data), but stale results after
  structural edits will confuse users.
- **Worker ↔ main thread serialization cost.** ViewPatch and
  DocViewModel are serialized as JSON across the `postMessage` boundary.
  For large initial loads, this could be expensive. May need binary
  encoding (e.g., MessagePack or a custom format) or `SharedArrayBuffer`
  for the view model.

---

## References

### CRDT Libraries

- yrs: https://docs.rs/yrs/latest/yrs/
- yrs sync: https://docs.rs/y-sync/latest/y_sync/
- automerge: https://docs.rs/automerge/latest/automerge/
- automerge sync: https://docs.rs/automerge/latest/automerge/sync/
- loro: https://docs.rs/loro/latest/loro/
- loro awareness: https://docs.rs/loro/latest/loro/awareness/
- diamond-types: https://docs.rs/diamond-types/latest/diamond_types/
- cola: https://docs.rs/cola_crdt/latest/cola_crdt/

### Y.js Ecosystem

- Yjs updates/sync: https://docs.yjs.dev/api/document-updates
- y-websocket: https://docs.yjs.dev/ecosystem/connection-provider/y-websocket
- y-indexeddb: https://docs.yjs.dev/ecosystem/database-provider/y-indexeddb
- y-prosemirror: https://docs.yjs.dev/ecosystem/editor-bindings/prosemirror

### Server Infrastructure

- axum WebSocket: https://docs.rs/axum/latest/axum/extract/ws/
- Hocuspocus: https://tiptap.dev/docs/hocuspocus/getting-started/overview
- Cloudflare DO WebSockets: https://developers.cloudflare.com/durable-objects/best-practices/websockets/
- Cloudflare DO limits: https://developers.cloudflare.com/durable-objects/platform/limits/
- partysocket: https://www.npmjs.com/package/partysocket
