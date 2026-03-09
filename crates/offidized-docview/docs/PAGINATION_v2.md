RFC: Word-Parity Pagination and Layout Engine for offidized-docedit/offidized-doclive

Status: Draft
Last updated: 2026-03-04
Implementation tracker: `PAGINATION_v2_TRACKER.md` (section-by-section compliance status)
Related specs: `PARITY_SPECS.md`, `EDITOR_PARITY.md`, `COLLABORATION.md`, `INTEROP.md`
Audience: offidized-docview/offidized-docx/offidized-docedit contributors
Primary goal: Deterministic, cross-browser-stable, Word-compatible pagination and layout (page breaks + line breaks + object placement) on the web.

⸻

0. Executive Summary

To achieve 1:1 pagination parity with Microsoft Word on the web, we must stop relying on the browser’s DOM layout engine for line breaking and pagination. Browser layout differs from Word’s layout in font fallback, shaping, justification, rounding, table width resolution, float wrapping, and footnote placement. “Render DOM, measure, and paginate” will always drift.

This RFC specifies a deterministic layout engine that produces a layout tree (pages → fragments → lines/boxes) plus a render display list (draw ops). The UI renders pages via canvas (or WebGPU) and uses DOM only for overlays (caret, selection, remote cursors). Editing uses a hidden textarea + custom caret/selection (or a “hybrid editing” mode initially), while the CRDT remains the only mutable source of truth.

This fits your current architecture:
• Rust/WASM (worker): CRDT, import/export, validation, intents.
• Worker JS + Canvas typography backend: paragraph shaping/line breaking/hit-testing + rendering.
• Main thread TS: input capture, overlays, persistence providers, scroll/viewport.

We also define a golden corpus + Word-PDF oracle testing pipeline, because Word parity is achieved by regression against real Word output.

⸻

1. Problem Statement

We need Word-accurate pagination in a web editor for .docx:
• Identical (or near-identical) line breaks and page breaks
• Correct handling of:
• Paragraph spacing/line rules, widow/orphan control, keep-with-next, keep-lines-together
• Tables (auto/fixed widths, row splitting, repeating headers, merged cells)
• Inline and floating images/shapes (wrap modes, anchors)
• Headers/footers, sections, columns
• Footnotes/endnotes (layout interaction with page height)
• Field codes / references (at least preserved; optionally evaluated)

Constraint: The CRDT is the single mutable model; layout is derived.

⸻

2. Definitions and Targets

2.1 “Word parity” target

We must choose one canonical Word engine as the oracle (build-time or deployment-time policy). Default recommendation:
• Canonical: Word Desktop on Windows (Microsoft 365 Current Channel)

Rationale: widest use; stable oracle; best-defined behaviors.

2.2 Success criteria (measurable)

A doc is “parity-pass” if all are true: 1. Same page count as Word 2. Same page break points (same first block/line on each page) 3. Pixel diff against Word-rendered PDF below threshold (configurable), excluding antialiasing noise 4. Stable output across Chrome/Edge/Safari/Firefox

⸻

3. Non-Goals (v1)

These are intentionally not required for first parity milestone:
• Editable tracked changes UI (but must preserve if present)
• Full fidelity for every DrawingML effect (3D, complex text-on-path)
• Full Word “layout compatibility flags” coverage on day one (but we must architect for it)
• Full accessibility parity with native DOM editors (we’ll provide an accessibility plan; see §15)

⸻

4. Architectural Overview

4.1 Core principle (unchanged)
• CRDT doc is the single mutable truth
• .docx import/export stays clone-and-patch via offidized-docx
• Layout is derived, deterministic, and cacheable

4.2 New major subsystem

Add a Layout + Render pipeline:

CRDT Doc (Rust/WASM)
└─> LayoutInputSnapshot (block structure + styles + inline tokens)
└─> Deterministic Layout Engine (Worker)
├─> LayoutTree (pages/fragments/lines/rects)
├─> HitTestIndex (coord<->(ParaId,UTF16 offset))
└─> DisplayList / Tile Render Plan
└─> Renderer (Canvas/WebGPU)

4.3 Worker split

The worker becomes a “document compute VM”:
• Rust/WASM: CRDT + doc model + intent application + export
• JS in worker: typography backend + rendering backend (CanvasKit or equivalent)

Main thread remains: input events, overlays, DOM composition, providers.

⸻

5. Technology Choice

5.1 Requirements for typography backend

We need:
• Deterministic glyph shaping and positioning
• BiDi + script runs
• Line breaking + justification
• Hit testing + selection rectangles
• Support for inline placeholders (your U+FFFC token model)

5.2 Recommended backend: CanvasKit + Paragraph

Use CanvasKit (Skia in WASM) as the cross-browser consistent renderer, and its paragraph module for text layout. This aligns extremely well with your design because SkParagraph uses U+FFFC placeholders for inline objects.

Note: This RFC does not require CanvasKit specifically, but it assumes we use a deterministic non-browser text layout engine with placeholder support. CanvasKit is the reference choice.

⸻

6. Data Model

6.1 Layout input snapshot

Derived from CRDT; immutable for a layout run.

type LayoutInputSnapshot = {
docVersion: u64, // monotonic counter from CRDT txn clock
sections: SectionInput[],
styles: ResolvedStyleRegistry, // resolved style inheritance
resources: {
images: Map<ImageRef, ImageMeta>, // bytes live in blob store
fonts: FontManifest,
}
}

6.2 Block model (layout-time)

type BlockInput =
| { kind: "paragraph", id: ParaId, props: ParaProps, richText: RichTextInput }
| { kind: "table", id: TableId, props: TableProps, grid: TableGridInput }
| { kind: "sectionBreak", id: SectionBreakId, props: SectionBreakProps }
| { kind: "floatingObjectAnchor", id: ObjId, anchor: AnchorInput, obj: FloatingObjInput };

6.3 RichTextInput (supports your token model)

type RichTextInput = {
textUtf16: string, // includes U+FFFC at token positions
spans: Array<{ start: u32, end: u32, attrs: TextAttrs }>, // marks
tokens: Array<{ index: u32, token: InlineToken }>, // index points to the U+FFFC
}

Invariant: every token maps to exactly one U+FFFC at the same UTF-16 offset; no “orphan sentinels”.

6.4 Layout output

type LayoutSnapshot = {
docVersion: u64,
pages: PageLayout[],
hitTest: HitTestIndex,
paint: DisplayListPlan, // either direct draw ops or tile plan
}

PageLayout

type PageLayout = {
pageIndex: u32,
paper: { w: i32, h: i32 }, // internal layout units
margins: { t: i32, r: i32, b: i32, l: i32 },
contentRect: Rect,
header?: StoryLayout,
footer?: StoryLayout,
footnotes?: FootnoteAreaLayout,
fragments: BlockFragment[],
}

BlockFragment (pagination result)

type BlockFragment =
| { kind: "paragraphFragment", paraId: ParaId, rect: Rect, lines: LineLayout[] }
| { kind: "tableFragment", tableId: TableId, rect: Rect, rows: TableRowFragment[] }
| { kind: "floatFragment", objId: ObjId, rect: Rect, wrap: WrapInfo }
| { kind: "sectionBreakFragment", rect: Rect };

⸻

7. Units, Rounding, and Determinism

7.1 Internal units

Pick one internal fixed-point coordinate system for all layout math. Recommended:
• Internal unit = 1/64 CSS px at 96 DPI (fixed-point int)

Why:
• Stable across zoom levels
• Fine enough to avoid rounding drift
• Efficient (integers)

Conversion examples:
• Points → px: px = pt \* 96 / 72
• Twips → pt: pt = twips / 20
• EMU → pt: pt = emu / 12700

7.2 Rounding contract

Define explicit rounding points:
• When resolving OOXML lengths to internal units: round half away from zero
• When distributing remaining width among columns: largest remainder method (deterministic)
• When snapping to device pixels for drawing: only at final render stage

Rule: layout tree coordinates never depend on canvas rasterization scale.

⸻

8. Fonts and Font Fallback (Parity prerequisite)

Word parity is impossible without font metric compatibility.

8.1 Font strategy

We support three tiers: 1. Exact fonts available (preferred): Calibri, Cambria, etc. 2. Metric-compatible substitutes (fallback): e.g., Carlito/Caladea-like families 3. Last-resort generic fallback: will break parity; still functional

8.2 Font manifest

From docx import + runtime availability:

type FontManifest = {
embeddedFonts?: Array<{ name: string, data: ArrayBuffer, style: "normal"|"italic", weight: u16 }>,
systemFallbackChains: Map<string, string[]>, // "Calibri" -> ["Calibri", "Carlito", "Arial", ...]
}

8.3 Deterministic fallback rule
• Always resolve font family with the same chain in all browsers
• Do not use browser CSS font fallback for layout; use the typography backend’s font manager

⸻

9. Paragraph Layout Specification

This is the “hard core” of parity.

9.1 Inputs
• Available width (page content width or table cell width)
• Paragraph props:
• indents: left/right/first-line/hanging
• alignment: left/right/center/justify
• line spacing: auto/exact/atLeast
• spacing before/after, contextual spacing
• tabs: stops, leaders
• bidi / rtl flags
• widow/orphan / keep constraints
• RichTextInput (text + marks + tokens)

9.2 Token/placeholder handling

Inline tokens are laid out as placeholders:
• Inline image token:
• placeholder width/height from token metadata (or derived from image meta)
• baseline alignment: bottom by default; support “center” alignment per OOXML if needed
• Tab token:
• zero-width placeholder in the paragraph string, but its advance is computed by tab algorithm (§9.6)
• LineBreak token:
• forced line break
• Preserved-but-not-editable tokens:
• generally occupy caret position but may be zero-width visually
• they still exist for hit-test mapping

9.3 Shaping and script runs
• Shape runs based on:
• font resolved
• script (Latin, Arabic, Han, etc.)
• direction (RTL/LTR)
• features (ligatures, etc.)
• Use backend’s shaping; ensure deterministic feature defaults.

9.4 Line breaking

Line breaking must consider:
• Unicode break opportunities
• NBSP and non-breaking segments
• Soft hyphen
• CJK kinsoku rules (v2 if needed; v1 can approximate but must be pluggable)

9.5 Justification
• Left/center/right: straightforward
• Justify: distribute extra advance:
• Latin: whitespace expansion
• CJK: inter-character expansion (when no spaces)
• Arabic: (future) kashida/spacing strategy hooks

Spec requirement: justification must be computed by the typography backend or by a deterministic wrapper around it; do not rely on canvas text rendering behavior alone.

9.6 Tabs

Tab behavior must match Word:

Inputs:
• tab stops (positions relative to paragraph start)
• default tab width
• leaders

Algorithm (per line): 1. Compute current x position from start of line + indentation. 2. Find next tab stop > x:
• if none: use default tab increments 3. Advance x to tab stop position 4. If leader enabled, draw leader glyphs between old x and new x

Tabs are not “characters” for line breaking; they are position jumps.

9.7 Line heights and spacing

Implement OOXML rules:
• lineRule="auto": line is multiplier based (e.g., w:line in 240ths of a line)
• exact: fixed line height; glyphs can overlap/clipped
• atLeast: minimum line height; grows if glyphs exceed

Paragraph spacing before/after:
• subject to contextualSpacing (“don’t add space between same style”)

9.8 Output from paragraph layout

type LineLayout = {
lineIndex: u32,
baselineY: i32,
ascent: i32,
descent: i32,
leading: i32,
xStart: i32,
xEnd: i32,
range: { startUtf16: u32, endUtf16: u32 },
runs: GlyphRunRef[], // reference into display list glyph runs
rectsForSelection?: Rect[], // optional cache
}

⸻

10. Pagination Algorithm Specification

Pagination is a vertical packing algorithm with constraints and split rules.

10.1 Page builder loop

For each section: 1. Establish page geometry (paper size, margins, header/footer regions) 2. Initialize current page with available height = contentRect.h 3. Flow blocks in order:
• measure block with available width
• if block fits: place it
• if block can split: split and place head; carry tail to next page
• else: start new page, place block there 4. Apply keep constraints and widow/orphan rules, potentially causing backtracking.

10.2 Keep and widow/orphan handling

Constraints:
• pageBreakBefore: forces new page before paragraph
• keepWithNext: paragraph must be on same page as the next block if possible
• keepLinesTogether: paragraph lines should not split across pages
• widow/orphan:
• ensure at least N lines at top/bottom (Word defaults: 2/2)

Backtracking strategy:
• When a constraint fails, roll back placement to a safe boundary (previous paragraph start) and repaginate from there with updated decision (e.g., move block to next page).

Important: Backtracking must be deterministic and bounded.

10.3 Paragraph splitting rules

Split paragraphs by line unless:
• keepLinesTogether (don’t split unless paragraph taller than a page)
• explicit “keep” rules force moving entire paragraph

When splitting:
• enforce widow/orphan minimum lines on both sides
• if impossible, relax rules in Word-compatible order (configurable)

10.4 Table splitting rules (high level)

Tables typically split by rows, with special rules in §11.

10.5 Columns (sections)

If section has columns:
• contentRect is split into N column rects
• flow blocks within a column; when column full, move to next column on same page; when last column full, new page

Column balancing (end of section) can be added later; v1 can do sequential fill.

⸻

11. Table Layout Specification

Tables are the biggest parity risk after typography.

11.1 Table width resolution

Inputs:
• table width type: auto/fixed/percent
• grid columns (tblGrid) widths
• cell widths
• cell margins
• borders

Algorithm v1 (deterministic, Word-like): 1. Initialize column widths from tblGrid if present; else equal split 2. For each cell:
• compute min-content width from laying out cell paragraphs at candidate widths (can be approximated with measuring “longest unbreakable segment”) 3. Solve for final widths:
• fixed columns keep their widths
• auto columns distribute remaining space proportionally
• clamp to min-content 4. Apply deterministic remainder distribution so sum(widths)=tableWidth

11.2 Cell layout

Each cell is its own flow:
• available width = columnWidth - cellPadding - borders
• lay out cell blocks
• compute cell content height
• apply vertical alignment

11.3 Row heights
• exact row height: clamp cell content; overflow is clipped
• atLeast: row grows to fit content
• auto: row height is max cell content height

11.4 Pagination for tables

Rules:
• repeating header rows: if row has tblHeader, replicate at top of each new page where table continues
• row splitting:
• if cantSplit or allowRowSplit=false: keep row together unless taller than page
• otherwise: allow splitting cell content across pages (advanced; v1 can defer and keep rows together)

v1 recommended behavior:
• Split tables only between rows
• If a single row > page height: allow overflow (like your current “no mid-paragraph split” limitation), but track as known limitation

⸻

12. Footnotes and Endnotes

Footnotes require iterative layout because they consume page height.

12.1 Iterative algorithm

For each page: 1. Layout main content assuming footnote area height = 0 2. Collect footnote refs that landed on this page 3. Layout footnote story (footnote paragraphs) into footnote area width 4. Compute footnote height and reduce main content height accordingly 5. Reflow the page; repeat until stable or max iterations (e.g., 3)

Stability criteria:
• same set of blocks on page
• same breakpoints in those blocks

⸻

13. Floating Objects (Images/Shapes) and Wrapping

For real Word parity, floating objects are required; for v1, define staged support.

13.1 v1 scope (recommended)
• Inline images: fully supported
• Floating images:
• square wrap
• anchor to paragraph (not page/margin yet)
• no tight/through polygon wrap initially

13.2 Wrap manager

Maintain active floats during line layout:
• Index floats by vertical span (yMin..yMax)
• For each line:
• compute excluded x intervals
• compute available segments
• choose segment to place text (left-to-right fill for LTR, mirrored for RTL)
• if no segment fits min line, break to next line (or move below float depending on wrap mode)

⸻

14. Rendering Plan

14.1 Two-stage output 1. Layout outputs DisplayListPlan 2. Renderer consumes plan and draws

14.2 Tiling (for large docs)
• Render pages into fixed-size tiles (e.g., 512×512)
• Cache tiles keyed by (docVersion, pageIndex, tileIndex, zoom)
• When layout changes:
• invalidate tiles intersecting changed fragments

14.3 Main thread composition

Main thread maintains:
• scroll container
• per-page canvas elements (or a virtualized canvas pool)
• DOM overlay layer for caret/selection/remote cursors

⸻

15. Editing Model on Top of Canvas Layout

Canvas rendering means no native DOM selection. Editing must be custom.

15.1 Input capture

Use:
• hidden <textarea> or <input> for keyboard + IME composition
• map input events to EditIntent and send to worker

15.2 Cursor and selection

Worker provides:
• caret rect for a position
• selection rects for a range

Main thread draws overlays.

15.3 Hit testing

On click: 1. determine page from scroll + event target 2. convert to page-local coords 3. call worker hitTest(x,y) → (paraId, utf16Offset, affinity) 4. call encode_position(paraId, utf16Offset) to get CRDT-relative cursor blob (still your canonical model)

15.4 Clipboard
• Copy:
• if selection spans tokens, serialize as HTML + plain text, mapping tokens to their visual representation (your existing policy)
• Paste:
• parse HTML/text
• generate intents (insertText, insertImage, etc.)
• enforce U+FFFC stripping at intent boundary (unchanged)

15.5 Accessibility plan (v1 and forward)

Canvas is not inherently accessible. Minimum plan:
• Maintain an offscreen DOM “accessibility tree” mirroring content (paragraph text only initially)
• Expose caret/selection via ARIA live region updates
• Provide keyboard navigation that matches layout

Full parity is larger; but we should not ship without at least basic screen reader support.

⸻

16. Worker and API Surface

16.1 Main ↔ Worker protocol (expanded)

Existing message types remain; add layout/paint messages.

// Main → Worker
type WorkerRequest =
| { type: "loadDocx"; data: ArrayBuffer }
| { type: "intent"; intent: EditIntent }
| { type: "applyUpdate"; update: ArrayBuffer }
| { type: "viewport"; viewport: ViewportState } // scroll/zoom/visible pages
| { type: "hitTest"; x: number; y: number; pageIndex: number }
| { type: "getSelectionRects"; anchor: Uint8Array; focus: Uint8Array }
| { type: "renderTiles"; pageIndices: number[]; zoom: number }
| { type: "saveDocx" }
| { type: "compact" };

// Worker → Main
type WorkerResponse =
| { type: "layout"; snapshot: LayoutSnapshotSummary } // lightweight; no big arrays
| { type: "tiles"; tiles: Array<{ pageIndex: number; tileIndex: number; bitmap: ImageBitmap }> }
| { type: "hitTestResult"; paraId: Uint8Array; utf16: number; affinity: "left"|"right" }
| { type: "selectionRects"; rects: Rect[] }
| { type: "localUpdate"; update: ArrayBuffer }
| { type: "savedDocx"; docx: ArrayBuffer }
| { type: "error"; message: string };

16.2 Rust/WASM FFI additions

Rust remains the CRDT authority. Add:
• export_layout_input_patch()
Returns changed paragraphs/tables (content + resolved styles + tokens) since last layout version.
• resolve_styles_snapshot()
Returns style registry + resolved computed styles (or a compact resolver table).
• get_block_structure()
Returns section/block ordering and properties.

JS layout engine calls these to avoid duplicating OOXML/style logic in JS.

⸻

17. Incremental Layout and Caching

17.1 Dirty tracking

From CRDT events:
• mark changed paragraphs/cells as dirty
• mark affected tables if cell changes affect row height
• mark section if page settings changed

17.2 Reflow window

Reflow from earliest dirty block forward until:
• page breaks stabilize for K pages in a row (e.g., K=2)
• or end of document

17.3 Caches
• Paragraph layout cache keyed by:
• paragraph id + version stamp
• available width
• resolved style hash
• Table cache keyed by:
• table id + version stamp
• available width
• Tile cache keyed by:
• (layoutVersion, pageIndex, tileIndex, zoom)

⸻

18. Validation and Security (Collaboration boundary)

All existing CRDT validation rules remain.

Additional layout-specific invariants:
• Tokens always map to U+FFFC positions
• No invalid UTF-16 boundaries (surrogate pairs) during hit testing
• Font loading must be sandboxed (limit size, validate format) if accepting user-supplied fonts

⸻

19. Test Plan and Parity Harness

19.1 Golden corpus

Build a corpus repo with categories:
• Typography: Latin + CJK + Arabic + mixed bidi
• Spacing: exact/atLeast/auto, contextual spacing
• Tabs: leaders and custom stops
• Numbering: multilevel lists
• Tables: merges, widths, repeating headers, row splits
• Images: inline and floating wrap modes
• Headers/footers: different first page, odd/even
• Sections: orientation changes, margins, columns
• Footnotes: heavy and mixed

19.2 Oracle generation (Word)

For each .docx:
• Generate Word PDF (CI step on Windows agent)
• Extract page images at fixed DPI
• Save metadata: page count, per-page text anchors (optional)

19.3 Diffing
• Render our pages to images at same DPI
• Pixel diff with tolerance
• Additionally compare page break mapping:
• first block id on each page
• first line’s paraId + utf16 offset

19.4 Performance tests
• 1000+ paragraph docs
• image-heavy docs
• table-heavy docs

Measure:
• keystroke → tile update latency
• memory footprint
• cache hit rate

⸻

20. Phased Implementation Roadmap

Phase 1 — Deterministic text layout + pagination (no tables/floats/footnotes)
• Integrate typography backend
• Paragraph layout + hit testing + selection rects
• Page building with keep/widow/orphan rules
• Canvas rendering (page canvases, no tiling initially)

Milestone: stable line/page breaks for text-only docs across browsers

Phase 2 — Tables
• Table grid + column width resolution
• Cell paragraph layout
• Table pagination by row, repeating headers
• Tile renderer + invalidation for performance

Phase 3 — Sections + headers/footers
• Story layout for headers/footers
• Section break handling, page settings changes
• Basic PAGE/NUMPAGES field evaluation (optional)

Phase 4 — Footnotes
• Iterative layout loop per page
• Footnote area rendering and splitting

Phase 5 — Floating objects + wrap
• Square wrap + paragraph anchoring
• Expand to tight/through polygons later

⸻

21. Risks and Mitigations

Risk: Font availability breaks parity

Mitigation:
• ship metric-compatible font pack
• embed fonts when available
• make “parity mode” warn if fonts missing

Risk: Complexity of full Word table model

Mitigation:
• incremental coverage
• large corpus + regression harness
• isolate table engine as a standalone module with dedicated tests

Risk: Canvas editing UX regression vs contenteditable

Mitigation:
• invest early in selection/IME correctness
• keep a “DOM editing” mode for internal/dev, not parity

⸻

22. Concrete Deliverables from This RFC
    1.  Layout snapshot schema (types in §6)
    2.  Worker protocol changes (§16)
    3.  Paragraph layout + pagination algorithms (§9–10)
    4.  Table + footnote + float algorithm specs (§11–13)
    5.  Parity harness plan (§19)
    6.  Phase roadmap (§20)

⸻

Appendix A: Minimal APIs to Implement First

If you want a very crisp “start here” list:

Worker internal modules
• LayoutCoordinator (JS): orchestrates layout, caches, tile invalidation
• Typography (JS): paragraph layout objects, hit test, rects for range
• Renderer (JS): tile/page drawing

Rust/WASM exports
• get_block_structure() -> binary
• get_layout_payload(blockIds[]) -> binary (text + spans + tokens + resolved styles)
• apply_intent(intent)
• encode_position/decode_position (already planned)

Main thread TS
• viewport reporting + tile display
• hidden textarea input + overlay caret/selection
• providers unchanged
