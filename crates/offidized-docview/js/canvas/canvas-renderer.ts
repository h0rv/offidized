// CanvasKit (Skia WASM) renderer implementing RendererAdapter.
//
// Converts DocViewModel into CanvasKit Paragraph objects, paginates them
// into pages, and draws each page onto a canvas via the CanvasPagePool.
// Supports editing: cursor rendering and selection
// highlighting are drawn as overlays on top of page content.

import type {
  RendererAdapter,
  HitTestResult,
  DocPosition,
  DocSelection,
  Rect,
  InlineImageHit,
  TableCellHit,
  TableCellPosition,
} from "../adapter.ts";
import type {
  DocViewModel,
  BodyItem,
  ParagraphModel,
  RunModel,
  TableModel,
  SectionModel,
  TableRowModel,
  FootnoteModel,
} from "../types.ts";
import { getCanvasKit } from "./canvaskit-loader.ts";
import { loadBundledFonts, resolveFontFamily } from "./font-manager.ts";
import { CanvasPagePool } from "./canvas-page-pool.ts";
import {
  type LayoutResult,
  type LayoutSnapshot,
  type MeasuredBlock,
  type ParagraphLineBox,
} from "./layout-engine.ts";
import { LayoutClient } from "./layout-client.ts";
import { extractRectList, getFirstRangeRect } from "./rect-utils.ts";
import {
  normalizeSelectionRange,
  selectionRangeInParagraphFragment,
} from "./selection-utils.ts";

// CanvasKit types are loaded at runtime from WASM. We use `any` to avoid
// a compile-time dependency on canvaskit-wasm type declarations.

/** Default page dimensions when no section is provided (US Letter). */
const DEFAULT_PAGE_WIDTH_PT = 612;
const DEFAULT_PAGE_HEIGHT_PT = 792;
const DEFAULT_MARGINS = { top: 72, right: 72, bottom: 72, left: 72 };

/** Default font size in points when a run has no explicit size. */
const DEFAULT_FONT_SIZE_PT = 11;

/** Estimated row height for tables when no explicit height is given. */
const TABLE_DEFAULT_ROW_HEIGHT_PT = 20;

/** Cell padding in points (matching Word's default 5.4pt horizontal). */
const TABLE_CELL_PADDING_PT = 5.4;

/** Fallback default tab stop width (0.5in, Word-like default). */
const DEFAULT_TAB_STOP_PT = 36;

/** Upper bound for deterministic layout refinement loops. */
const MAX_LAYOUT_REFINEMENT_PASSES = 3;

/** Reserved footnote separator spacing above footnote content. */
const FOOTNOTE_SEPARATOR_PT = 8;

/** Gap between footnotes in the reserved area. */
const FOOTNOTE_GAP_PT = 4;

/** Simple cap to avoid footnotes consuming the full content area in v1. */
const FOOTNOTE_MAX_AREA_RATIO = 0.45;

/** Cursor width in points (1pt ~= 1.33 CSS px). */
const CURSOR_WIDTH_PT = 1;
/** Minimum visual caret height in CSS px to avoid dot-caret collapse. */
const CURSOR_MIN_HEIGHT_PX = 12;

/** Selection highlight color: semi-transparent blue matching ::selection. */
const SELECTION_COLOR_RGBA = [194, 220, 255, 0.72] as const;

/** Clip padding to avoid trimming glyph side bearings/ascenders. */
const PARAGRAPH_CLIP_PAD_PT = 2;

/** Info tracked per laid-out paragraph for hit testing. */
interface LaidOutParagraph {
  /** CanvasKit Paragraph object (has getGlyphPositionAtCoordinate). */
  paragraph: any;
  /** Page index this paragraph was placed on. */
  pageIndex: number;
  /** Y offset within the page content area, in points. */
  yPt: number;
  /** X offset within the page content area, in points. */
  xPt: number;
  /** Width used for layout, in points. */
  widthPt: number;
  /** Height after layout, in points. */
  heightPt: number;
  /** Index into the original model.body array. */
  bodyIndex: number;
  /** Top y within the paragraph source content for this fragment. */
  sourceTopPt: number;
  /** UTF-16 range covered by this fragment. */
  utf16Start: number;
  utf16End: number;
  /** Visual list-marker length present in the laid-out paragraph. */
  markerUtf16Len: number;
  /** Line index range covered by this fragment. */
  lineStart: number;
  lineEnd: number;
}

/** Info tracked per laid-out table for hit testing. */
interface LaidOutTable {
  pageIndex: number;
  yPt: number;
  xPt: number;
  widthPt: number;
  heightPt: number;
  bodyIndex: number;
  rowStart: number;
  rowEnd: number;
  repeatedHeaderRowCount?: number;
}

/** Internal page representation during pagination. */
interface PageInfo {
  pageIndex: number;
  sectionIndex: number;
  section: SectionModel;
  items: Array<
    | { type: "paragraph"; ref: LaidOutParagraph }
    | { type: "table"; ref: LaidOutTable }
  >;
}

interface MeasuredParagraphEntry {
  bodyIndex: number;
  item: Extract<BodyItem, { type: "paragraph" }>;
  section: SectionModel;
  sectionIndex: number;
  paragraph: any;
  lineBoxes: ParagraphLineBox[];
  markerUtf16Len: number;
}

interface MeasuredTableEntry {
  bodyIndex: number;
  item: Extract<BodyItem, { type: "table" }>;
  section: SectionModel;
  sectionIndex: number;
  tableInfo: {
    rows: Array<{
      heightPt: number;
      isHeader?: boolean;
      repeatHeader?: boolean;
      keepTogether?: boolean;
      keepWithNext?: boolean;
      cantSplit?: boolean;
    }>;
    totalHeightPt: number;
  };
}

type MeasuredEntry = MeasuredParagraphEntry | MeasuredTableEntry;

interface ParagraphMeasurementCacheEntry {
  key: string;
  paragraph: any;
  lineBoxes: ParagraphLineBox[];
  markerUtf16Len: number;
}

interface TableMeasurementCacheEntry {
  key: string;
  tableInfo: {
    rows: Array<{
      heightPt: number;
      isHeader?: boolean;
      repeatHeader?: boolean;
      keepTogether?: boolean;
      keepWithNext?: boolean;
      cantSplit?: boolean;
    }>;
    totalHeightPt: number;
  };
}

type ParagraphToken =
  | {
      kind: "text";
      text: string;
      style: any;
      utf16Len: number;
    }
  | {
      kind: "placeholder";
      role: "inlineImage" | "tab";
      widthPt: number;
      heightPt: number;
      alignment: any;
      baseline: any;
      baselineOffsetPt: number;
      tabIndex?: number;
    };

/**
 * CanvasKit-based renderer for DocViewModel with editing support.
 *
 * Create with `CanvasRenderer.create(container)` which handles async
 * initialization of CanvasKit and default fonts. Supports cursor
 * rendering with caret and selection highlighting.
 */
export class CanvasRenderer implements RendererAdapter {
  private ck: any;
  private fontCollection: any;
  private scrollContainer: HTMLElement;
  private wrapper: HTMLElement;
  private pool: CanvasPagePool;
  private layoutClient: LayoutClient;
  private dummyInput: HTMLDivElement;
  private caretOverlay: HTMLDivElement;

  // Layout state from last renderModel call
  private laidOutParagraphs: LaidOutParagraph[] = [];
  private laidOutTables: LaidOutTable[] = [];
  private pages: PageInfo[] = [];
  private layoutSnapshot: LayoutSnapshot | null = null;
  private layoutVersion = 0;
  private currentModel: DocViewModel | null = null;
  private imageCache: Map<number, any> = new Map();
  private scrollListener: (() => void) | null = null;
  private pendingDrawPages = new Set<number>();
  private pageDrawRafHandles = new Map<number, number>();
  private visibleDrawRaf = 0;
  private renderEpoch = 0;
  private paragraphPagesByBodyIndex = new Map<number, Set<number>>();
  private paragraphsByBodyIndex = new Map<number, LaidOutParagraph[]>();
  private paragraphCharCountByBodyIndex = new Map<number, number>();
  private paragraphMeasureCache = new Map<
    number,
    ParagraphMeasurementCacheEntry
  >();
  private tableMeasureCache = new Map<number, TableMeasurementCacheEntry>();

  // Editing state: cursor and selection
  private cursorPos: DocPosition | null = null;
  private selection: DocSelection | null = null;
  private cursorVisible = false;
  private lastCaretOverlayRect: {
    left: number;
    top: number;
    width: number;
    height: number;
  } | null = null;

  private constructor(
    ck: any,
    fontCollection: any,
    scrollContainer: HTMLElement,
    wrapper: HTMLElement,
  ) {
    this.ck = ck;
    this.fontCollection = fontCollection;
    this.scrollContainer = scrollContainer;
    this.wrapper = wrapper;
    this.pool = new CanvasPagePool(scrollContainer, wrapper);
    this.layoutClient = new LayoutClient();

    // Fallback focus target used when the external input adapter has not
    // focused its own input element yet.
    this.dummyInput = document.createElement("div");
    this.dummyInput.tabIndex = 0;
    this.dummyInput.style.position = "absolute";
    this.dummyInput.style.opacity = "0";
    this.dummyInput.style.pointerEvents = "none";
    this.dummyInput.style.width = "0";
    this.dummyInput.style.height = "0";
    this.dummyInput.style.overflow = "hidden";
    scrollContainer.appendChild(this.dummyInput);

    const caretOverlay = document.createElement("div");
    caretOverlay.dataset.docviewCaretOverlay = "1";
    caretOverlay.style.position = "absolute";
    caretOverlay.style.pointerEvents = "none";
    caretOverlay.style.zIndex = "5";
    caretOverlay.style.background = "rgb(26, 115, 232)";
    caretOverlay.style.display = "none";
    this.caretOverlay = caretOverlay;
    scrollContainer.appendChild(caretOverlay);
  }

  /**
   * Async factory. Loads CanvasKit WASM and the default font, then
   * returns a ready-to-use renderer.
   *
   * @param container - The scroll container element. A wrapper div is
   *   created inside it to hold page canvases.
   */
  static async create(container: HTMLElement): Promise<CanvasRenderer> {
    const ck = (await getCanvasKit()) as any;
    const bundledFonts = await loadBundledFonts();

    // Build a FontCollection with a TypefaceFontProvider
    const provider = ck.TypefaceFontProvider.Make();
    for (const font of bundledFonts) {
      provider.registerFont(font.data, font.family);
    }
    const fontCollection = ck.FontCollection.Make();
    fontCollection.setDefaultFontManager(provider);
    fontCollection.enableFontFallback();

    const wrapper = document.createElement("div");
    wrapper.className = "docview-paginated";
    if (getComputedStyle(container).position === "static") {
      container.style.position = "relative";
    }
    container.appendChild(wrapper);

    return new CanvasRenderer(ck, fontCollection, container, wrapper);
  }

  // ---------- RendererAdapter: renderModel ----------

  renderModel(model: DocViewModel): void {
    this.cancelPendingDraws();
    this.renderEpoch += 1;
    this.currentModel = model;
    this.rebuildParagraphCharCountCache(model);
    this.disposeImageCache();

    const defaultSection: SectionModel = model.sections[0] ?? {
      pageWidthPt: DEFAULT_PAGE_WIDTH_PT,
      pageHeightPt: DEFAULT_PAGE_HEIGHT_PT,
      orientation: "portrait",
      margins: DEFAULT_MARGINS,
    };

    // Step 1: Build paragraphs and collect deterministic line metrics.
    const measured = this.measureBodyItems(model, defaultSection);

    // Step 2: Convert measured data to layout-engine block inputs.
    const sectionsForLayout =
      model.sections.length > 0 ? model.sections : [defaultSection];
    const blocks = this.buildMeasuredBlocks(measured);

    // Step 3: Deterministic pagination and block-fragment generation.
    this.layoutVersion += 1;
    try {
      const layout = this.layoutWithFootnoteReservationSync(
        blocks,
        sectionsForLayout,
        model,
        this.layoutVersion,
      );
      if (this.currentModel !== model) return;
      this.applyLayoutResult(
        layout,
        measured,
        sectionsForLayout,
        defaultSection,
      );
    } catch (err) {
      console.error("layout failed:", err);
      this.updateCaretOverlay();
    }
  }

  private buildMeasuredBlocks(measured: MeasuredEntry[]): MeasuredBlock[] {
    return measured.map((entry) => {
      if ("paragraph" in entry) {
        if (entry.item.type !== "paragraph") {
          throw new Error("invalid measured paragraph entry");
        }
        const para = entry.item;
        return {
          kind: "paragraph",
          bodyIndex: entry.bodyIndex,
          sectionIndex: entry.sectionIndex,
          pageBreakBefore: para.pageBreakBefore ?? false,
          keepNext: para.keepNext ?? false,
          keepLines: para.keepLines ?? false,
          spacingBeforePt: para.spacingBeforePt ?? 0,
          spacingAfterPt: para.spacingAfterPt ?? 8,
          indentLeftPt: para.indents?.leftPt ?? 0,
          indentRightPt: para.indents?.rightPt ?? 0,
          lines: entry.lineBoxes,
          paragraphWidthPt: entry.paragraph.getMaxWidth?.() ?? 0,
        };
      }

      return {
        kind: "table",
        bodyIndex: entry.bodyIndex,
        sectionIndex: entry.sectionIndex,
        rows: entry.tableInfo.rows,
        totalHeightPt: entry.tableInfo.totalHeightPt,
      };
    });
  }

  private layoutWithFootnoteReservationSync(
    blocks: MeasuredBlock[],
    sections: SectionModel[],
    model: DocViewModel,
    docVersion: number,
  ): LayoutResult {
    let reservedBottomByPagePt: number[] = [];
    let layout: LayoutResult | null = null;

    for (let pass = 0; pass < MAX_LAYOUT_REFINEMENT_PASSES; pass++) {
      layout = this.layoutClient.layoutSync({
        blocks,
        sections,
        docVersion,
        config:
          reservedBottomByPagePt.length > 0
            ? { reservedBottomByPagePt }
            : undefined,
      });
      const nextReserved = this.computeFootnoteReservations(
        model,
        sections,
        layout,
      );
      if (this.sameNumberArray(nextReserved, reservedBottomByPagePt)) {
        return layout;
      }
      reservedBottomByPagePt = nextReserved;
    }

    return (
      layout ?? this.layoutClient.layoutSync({ blocks, sections, docVersion })
    );
  }

  private computeFootnoteReservations(
    model: DocViewModel,
    sections: SectionModel[],
    layout: LayoutResult,
  ): number[] {
    if (model.footnotes.length === 0 || layout.pages.length === 0) {
      return [];
    }

    const refsByPage = this.collectFootnoteRefsByPage(model, layout);
    const footnotesById = new Map<number, FootnoteModel>();
    for (const fn of model.footnotes) {
      footnotesById.set(fn.id, fn);
    }

    const reserved: number[] = [];
    for (const page of layout.pages) {
      const section = sections[page.sectionIndex];
      if (!section) {
        reserved.push(0);
        continue;
      }
      const refs = refsByPage.get(page.pageIndex);
      if (!refs || refs.size === 0) {
        reserved.push(0);
        continue;
      }
      reserved.push(this.measureFootnoteAreaPt(section, refs, footnotesById));
    }
    return reserved;
  }

  private collectFootnoteRefsByPage(
    model: DocViewModel,
    layout: LayoutResult,
  ): Map<number, Set<number>> {
    const refsByPage = new Map<number, Set<number>>();
    const paragraphRefSpans = new Map<
      number,
      Array<{ startUtf16: number; endUtf16: number; footnoteId: number }>
    >();

    for (const page of layout.pages) {
      for (const fragment of page.fragments) {
        if (fragment.kind !== "paragraphFragment") continue;
        const bodyItem = model.body[fragment.bodyIndex];
        if (!bodyItem || bodyItem.type !== "paragraph") continue;

        let spans = paragraphRefSpans.get(fragment.bodyIndex);
        if (!spans) {
          spans = this.extractParagraphFootnoteSpans(bodyItem);
          paragraphRefSpans.set(fragment.bodyIndex, spans);
        }
        if (spans.length === 0) continue;

        const pageSet = refsByPage.get(page.pageIndex) ?? new Set<number>();
        for (const span of spans) {
          if (
            span.endUtf16 > fragment.utf16Start &&
            span.startUtf16 < fragment.utf16End
          ) {
            pageSet.add(span.footnoteId);
          }
        }
        if (pageSet.size > 0) {
          refsByPage.set(page.pageIndex, pageSet);
        }
      }
    }

    return refsByPage;
  }

  private extractParagraphFootnoteSpans(
    paragraph: ParagraphModel,
  ): Array<{ startUtf16: number; endUtf16: number; footnoteId: number }> {
    const spans: Array<{
      startUtf16: number;
      endUtf16: number;
      footnoteId: number;
    }> = [];
    let offset = 0;

    for (const run of paragraph.runs) {
      if (run.inlineImage) {
        offset += 1;
        continue;
      }
      if (run.hasBreak) {
        offset += (run.text?.length ?? 0) + 1;
        continue;
      }
      if (run.hasTab) {
        offset += (run.text?.length ?? 0) + 1;
        continue;
      }
      if (run.footnoteRef != null) {
        const refText = String(run.footnoteRef);
        if (refText.length > 0) {
          spans.push({
            startUtf16: offset,
            endUtf16: offset + refText.length,
            footnoteId: run.footnoteRef,
          });
        }
        offset += refText.length;
        continue;
      }
      if (run.endnoteRef != null) {
        offset += String(run.endnoteRef).length;
        continue;
      }
      offset += run.text?.length ?? 0;
    }

    return spans;
  }

  private measureFootnoteAreaPt(
    section: SectionModel,
    refs: Set<number>,
    footnotesById: Map<number, FootnoteModel>,
  ): number {
    const widthPt = Math.max(
      40,
      section.pageWidthPt - section.margins.left - section.margins.right,
    );
    const contentHeightPt = Math.max(
      0,
      section.pageHeightPt - section.margins.top - section.margins.bottom,
    );
    const maxAreaPt = contentHeightPt * FOOTNOTE_MAX_AREA_RATIO;

    const sortedRefs = Array.from(refs).sort((a, b) => a - b);
    if (sortedRefs.length === 0) return 0;

    let total = FOOTNOTE_SEPARATOR_PT;
    const ck = this.ck;
    for (const refId of sortedRefs) {
      const model = footnotesById.get(refId);
      const text = model?.text?.trim() ?? "";
      const paraStyle = new ck.ParagraphStyle({
        textAlign: ck.TextAlign.Left,
        textDirection: ck.TextDirection.LTR,
        textStyle: { color: ck.Color(0, 0, 0, 1) },
      });
      const builder = ck.ParagraphBuilder.MakeFromFontCollection(
        paraStyle,
        this.fontCollection,
      );
      const style = new ck.TextStyle({
        fontFamilies: [resolveFontFamily()],
        fontSize: Math.max(8, DEFAULT_FONT_SIZE_PT * 0.85),
        color: ck.Color(0, 0, 0, 1),
      });
      builder.pushStyle(style);
      builder.addText(`${refId}. ${text}`);
      builder.pop();
      const paragraph = builder.build();
      paragraph.layout(widthPt);
      total += Math.max(DEFAULT_FONT_SIZE_PT, paragraph.getHeight?.() ?? 0);
      total += FOOTNOTE_GAP_PT;
      paragraph.delete();
    }

    if (maxAreaPt <= 0) return 0;
    return Math.max(0, Math.min(total, maxAreaPt));
  }

  private sameNumberArray(a: number[], b: number[]): boolean {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) {
      if (Math.abs((a[i] ?? 0) - (b[i] ?? 0)) > 1e-3) return false;
    }
    return true;
  }

  private applyLayoutResult(
    layout: LayoutResult,
    measured: MeasuredEntry[],
    sectionsForLayout: SectionModel[],
    defaultSection: SectionModel,
  ): void {
    this.laidOutParagraphs = [];
    this.laidOutTables = [];
    this.pages = [];
    this.paragraphPagesByBodyIndex.clear();
    this.paragraphsByBodyIndex.clear();
    this.layoutSnapshot = layout.snapshot;
    this.pages = layout.pages.map((page) => ({
      pageIndex: page.pageIndex,
      sectionIndex: page.sectionIndex,
      section: sectionsForLayout[page.sectionIndex] ?? defaultSection,
      items: [],
    }));

    // Build lookup for paragraph objects by body index.
    const paragraphMap = new Map<number, MeasuredParagraphEntry>();
    for (const entry of measured) {
      if ("paragraph" in entry) {
        paragraphMap.set(entry.bodyIndex, entry);
      }
    }

    // Translate layout fragments to renderer placements.
    for (const page of layout.pages) {
      const pageInfo = this.pages[page.pageIndex];
      if (!pageInfo) continue;
      for (const fragment of page.fragments) {
        if (fragment.kind === "paragraphFragment") {
          const measuredPara = paragraphMap.get(fragment.bodyIndex);
          if (!measuredPara) continue;
          const lp: LaidOutParagraph = {
            paragraph: measuredPara.paragraph,
            pageIndex: page.pageIndex,
            yPt: fragment.yPt,
            xPt: fragment.xPt,
            widthPt: fragment.wPt,
            heightPt: fragment.hPt,
            bodyIndex: fragment.bodyIndex,
            sourceTopPt: fragment.sourceTopPt,
            utf16Start: fragment.utf16Start,
            utf16End: fragment.utf16End,
            markerUtf16Len: measuredPara.markerUtf16Len,
            lineStart: fragment.lineStart,
            lineEnd: fragment.lineEnd,
          };
          this.laidOutParagraphs.push(lp);
          const bodyFragments =
            this.paragraphsByBodyIndex.get(lp.bodyIndex) ?? [];
          bodyFragments.push(lp);
          this.paragraphsByBodyIndex.set(lp.bodyIndex, bodyFragments);
          const pagesForBody =
            this.paragraphPagesByBodyIndex.get(lp.bodyIndex) ??
            new Set<number>();
          pagesForBody.add(lp.pageIndex);
          this.paragraphPagesByBodyIndex.set(lp.bodyIndex, pagesForBody);
          pageInfo.items.push({ type: "paragraph", ref: lp });
        } else {
          const lt: LaidOutTable = {
            pageIndex: page.pageIndex,
            yPt: fragment.yPt,
            xPt: fragment.xPt,
            widthPt: fragment.wPt,
            heightPt: fragment.hPt,
            bodyIndex: fragment.bodyIndex,
            rowStart: fragment.rowStart,
            rowEnd: fragment.rowEnd,
            repeatedHeaderRowCount: fragment.repeatedHeaderRowCount,
          };
          this.laidOutTables.push(lt);
          pageInfo.items.push({ type: "table", ref: lt });
        }
      }
    }
    for (const fragments of this.paragraphsByBodyIndex.values()) {
      fragments.sort((a, b) => a.pageIndex - b.pageIndex || a.yPt - b.yPt);
    }

    // Step 4: Set up page pool
    const pageWidths = this.pages.map((p) => p.section.pageWidthPt);
    const pageHeights = this.pages.map((p) => p.section.pageHeightPt);
    this.pool.setPageCount(this.pages.length, pageWidths, pageHeights);

    // Step 5: Initial draw of visible pages
    this.pool.updateVisibility();
    this.drawVisiblePages();
    this.updateCaretOverlay();

    // Step 6: Listen for scroll to virtualize
    if (this.scrollListener) {
      this.scrollContainer.removeEventListener("scroll", this.scrollListener);
      this.scrollListener = null;
    }
    this.scrollListener = () => {
      this.updateCaretOverlay();
      if (this.visibleDrawRaf !== 0) return;
      const epoch = this.renderEpoch;
      this.visibleDrawRaf = requestAnimationFrame(() => {
        this.visibleDrawRaf = 0;
        if (epoch !== this.renderEpoch) return;
        this.pool.updateVisibility();
        this.drawVisiblePages();
        this.updateCaretOverlay();
      });
    };
    this.scrollContainer.addEventListener("scroll", this.scrollListener, {
      passive: true,
    });
  }

  // ---------- RendererAdapter: hitTest ----------

  hitTest(x: number, y: number): HitTestResult | null {
    if (this.pages.length === 0) return null;

    // Determine which page was clicked by checking canvas positions
    for (let i = 0; i < this.pages.length; i++) {
      const canvas = this.pool.getCanvas(i);
      if (!canvas) continue;
      const rect = canvas.getBoundingClientRect();
      if (
        x >= rect.left &&
        x <= rect.right &&
        y >= rect.top &&
        y <= rect.bottom
      ) {
        // Convert screen pixels to page-local points.
        // Canvas CSS size is pageWidthPt "pt", rendered by the browser.
        // We need screen px → pt conversion via the bounding rect.
        const page = this.pages[i];
        if (!page) continue;
        const screenToPtX = page.section.pageWidthPt / rect.width;
        const screenToPtY = page.section.pageHeightPt / rect.height;
        const localXPt = (x - rect.left) * screenToPtX;
        const localYPt = (y - rect.top) * screenToPtY;

        return this.hitTestPage(i, localXPt, localYPt);
      }
    }

    return null;
  }

  // ---------- RendererAdapter: editing methods ----------

  setCursor(pos: DocPosition): void {
    const previousSelectionPages = this.getSelectionPageIndices(this.selection);
    const normalizedPos = this.normalizeCursorPosition(pos);
    this.cursorPos = normalizedPos;
    this.selection = null;
    this.showCursor();
    // Redraw pages that previously showed selection highlights.
    for (const pageIndex of previousSelectionPages) {
      this.requestPageDraw(pageIndex);
    }
  }

  setSelection(sel: DocSelection): void {
    const previousPages = this.getSelectionPageIndices(this.selection);
    this.selection = sel;
    this.cursorPos = null;
    this.hideCursor();
    const nextPages = this.getSelectionPageIndices(sel);

    const pagesToRedraw = new Set<number>();
    for (const pageIndex of previousPages) pagesToRedraw.add(pageIndex);
    for (const pageIndex of nextPages) pagesToRedraw.add(pageIndex);

    for (const pageIndex of pagesToRedraw) {
      this.requestPageDraw(pageIndex);
    }
  }

  getSelection(): DocSelection | null {
    if (this.selection) return this.selection;
    if (this.cursorPos) {
      // Collapsed selection: anchor == focus
      return {
        anchor: { ...this.cursorPos },
        focus: { ...this.cursorPos },
      };
    }
    return null;
  }

  getCursorRect(): Rect | null {
    const pos = this.cursorPos ?? this.selection?.focus;
    if (!pos) return null;
    return this.getCursorRectForPosition(pos);
  }

  getCursorRectForPosition(pos: DocPosition): Rect | null {
    if (!pos) return null;

    const lp = this.findLaidOutParagraphForPosition(pos);
    if (!lp) return null;

    const canvas = this.pool.getCanvas(lp.pageIndex);
    if (!canvas) return null;

    const page = this.pages[lp.pageIndex];
    if (!page) return null;

    const section = page.section;
    const caret = this.getCaretRectInParagraph(lp, pos.charOffset);
    if (!caret) return null;

    // Position in points relative to the page origin
    const pagePtX = section.margins.left + lp.xPt + caret.xPt;
    const pagePtY = section.margins.top + lp.yPt + caret.yPt - lp.sourceTopPt;

    // Convert from page-local points to screen pixels using the canvas bounding rect.
    // The canvas CSS size is pageWidthPt "pt" = pageWidthPt * (96/72) CSS px.
    // The canvas bounding rect gives us the actual rendered size on screen.
    const canvasRect = canvas.getBoundingClientRect();
    const ptToScreenX = canvasRect.width / section.pageWidthPt;
    const ptToScreenY = canvasRect.height / section.pageHeightPt;

    return {
      x: canvasRect.left + pagePtX * ptToScreenX,
      y: canvasRect.top + pagePtY * ptToScreenY,
      w: CURSOR_WIDTH_PT * ptToScreenX,
      h: caret.heightPt * ptToScreenY,
    };
  }

  getSelectionRects(sel: DocSelection): Rect[] {
    const normalized = normalizeSelectionRange(sel);
    if (!normalized) return [];
    const rects: Rect[] = [];

    for (const lp of this.laidOutParagraphs) {
      const paragraphEnd = this.getParagraphCharCountByBodyIndex(lp.bodyIndex);
      const range = selectionRangeInParagraphFragment(
        normalized,
        lp.bodyIndex,
        lp.utf16Start,
        lp.utf16End,
        paragraphEnd,
      );
      if (!range) continue;

      const canvas = this.pool.getCanvas(lp.pageIndex);
      const page = this.pages[lp.pageIndex];
      if (!canvas || !page) continue;

      const canvasRect = canvas.getBoundingClientRect();
      const ptToScreenX = canvasRect.width / page.section.pageWidthPt;
      const ptToScreenY = canvasRect.height / page.section.pageHeightPt;
      const originX = page.section.margins.left * ptToScreenX;
      const originY = page.section.margins.top * ptToScreenY;

      const rawRects = lp.paragraph.getRectsForRange(
        range.start + lp.markerUtf16Len,
        range.end + lp.markerUtf16Len,
        this.ck.RectHeightStyle.Max,
        this.ck.RectWidthStyle.Tight,
      );
      for (const rect of extractRectList(rawRects as unknown)) {
        const rx = canvasRect.left + originX + (lp.xPt + rect[0]) * ptToScreenX;
        const ry =
          canvasRect.top +
          originY +
          (lp.yPt + rect[1] - lp.sourceTopPt) * ptToScreenY;
        const rw = (rect[2] - rect[0]) * ptToScreenX;
        const rh = (rect[3] - rect[1]) * ptToScreenY;
        if (!(rw > 0 && rh > 0)) continue;
        rects.push({ x: rx, y: ry, w: rw, h: rh });
      }
    }

    return rects;
  }

  getInlineImageAtPoint(x: number, y: number): InlineImageHit | null {
    if (!this.currentModel || this.pages.length === 0) return null;
    for (const lp of this.laidOutParagraphs) {
      const hit = this.hitTestInlineImageInParagraph(lp, x, y);
      if (hit) return hit;
    }
    return null;
  }

  getInlineImageRect(pos: DocPosition): Rect | null {
    if (!this.currentModel || this.pages.length === 0) return null;
    for (const lp of this.laidOutParagraphs) {
      if (lp.bodyIndex !== pos.bodyIndex) continue;
      const hit = this.inlineImageRectInParagraph(lp, pos.charOffset);
      if (hit) return hit.rect;
    }
    return null;
  }

  getTableCellAtPoint(x: number, y: number): TableCellHit | null {
    if (!this.currentModel || this.pages.length === 0) return null;
    for (let pageIndex = 0; pageIndex < this.pages.length; pageIndex += 1) {
      const canvas = this.pool.getCanvas(pageIndex);
      const page = this.pages[pageIndex];
      if (!canvas || !page) continue;
      const rect = canvas.getBoundingClientRect();
      if (x < rect.left || x > rect.right || y < rect.top || y > rect.bottom) {
        continue;
      }
      const screenToPtX = page.section.pageWidthPt / rect.width;
      const screenToPtY = page.section.pageHeightPt / rect.height;
      const localXPt = (x - rect.left) * screenToPtX;
      const localYPt = (y - rect.top) * screenToPtY;
      const hit = this.hitTestTableCell(pageIndex, localXPt, localYPt);
      if (!hit) return null;
      const cellRect = this.getTableCellRect(hit);
      if (!cellRect) return null;
      return { ...hit, rect: cellRect };
    }
    return null;
  }

  getTableCellRect(cell: TableCellPosition): Rect | null {
    if (!this.currentModel) return null;
    const bodyItem = this.currentModel.body[cell.bodyIndex];
    if (!bodyItem || bodyItem.type !== "table") return null;
    for (const lt of this.laidOutTables) {
      if (lt.bodyIndex !== cell.bodyIndex) continue;
      if (cell.row < lt.rowStart || cell.row >= lt.rowEnd) continue;
      const canvas = this.pool.getCanvas(lt.pageIndex);
      const page = this.pages[lt.pageIndex];
      if (!canvas || !page) continue;
      const rect = canvas.getBoundingClientRect();
      const ptToScreenX = rect.width / page.section.pageWidthPt;
      const ptToScreenY = rect.height / page.section.pageHeightPt;
      const tableRect = this.tableCellRectInPagePt(bodyItem, lt, page, cell);
      if (!tableRect) return null;
      return {
        x: rect.left + tableRect.xPt * ptToScreenX,
        y: rect.top + tableRect.yPt * ptToScreenY,
        w: tableRect.wPt * ptToScreenX,
        h: tableRect.hPt * ptToScreenY,
      };
    }
    return null;
  }

  getInputElement(): HTMLElement {
    return this.dummyInput;
  }

  isFocused(): boolean {
    return this.hasEditorFocus();
  }

  focus(): void {
    if (!this.hasEditorFocus()) {
      const target = this.getPreferredFocusTarget();
      target?.focus({ preventScroll: true });
    }
    if (this.cursorPos) {
      this.showCursor();
    }
  }

  getScrollContainer(): HTMLElement {
    return this.scrollContainer;
  }

  destroy(): void {
    this.hideCursor();
    this.cancelPendingDraws();
    if (this.scrollListener) {
      this.scrollContainer.removeEventListener("scroll", this.scrollListener);
      this.scrollListener = null;
    }
    this.layoutClient.destroy();
    this.pool.destroy();
    this.disposeImageCache();
    if (this.dummyInput.parentNode) {
      this.dummyInput.parentNode.removeChild(this.dummyInput);
    }
    if (this.caretOverlay.parentNode) {
      this.caretOverlay.parentNode.removeChild(this.caretOverlay);
    }
    this.wrapper.innerHTML = "";
    this.laidOutParagraphs = [];
    this.laidOutTables = [];
    this.pages = [];
    this.layoutSnapshot = null;
    this.currentModel = null;
    this.cursorPos = null;
    this.selection = null;
    this.cursorVisible = false;
    this.paragraphPagesByBodyIndex.clear();
    this.paragraphsByBodyIndex.clear();
    this.paragraphCharCountByBodyIndex.clear();
    this.lastCaretOverlayRect = null;
    this.clearMeasurementCaches();
  }

  // ==========================================================================
  // Internal: measurement and paragraph building
  // ==========================================================================

  /**
   * For each body item, build paragraph metrics (line boxes) or table metrics.
   */
  private measureBodyItems(
    model: DocViewModel,
    defaultSection: SectionModel,
  ): MeasuredEntry[] {
    const results: MeasuredEntry[] = [];
    const seenParagraphIndices = new Set<number>();
    const seenTableIndices = new Set<number>();

    for (let i = 0; i < model.body.length; i++) {
      const item = model.body[i]!;
      const sectionIdx = item.sectionIndex;
      const section = model.sections[sectionIdx] ?? defaultSection;
      const columnWidth = this.sectionColumnWidth(section);

      if (item.type === "paragraph") {
        seenParagraphIndices.add(i);
        const cacheKey = this.paragraphMeasureCacheKey(
          item,
          section,
          columnWidth,
        );
        const cached = this.paragraphMeasureCache.get(i);
        let para: any;
        let lineBoxes: ParagraphLineBox[];
        let markerUtf16Len: number;
        if (cached && cached.key === cacheKey) {
          para = cached.paragraph;
          lineBoxes = cached.lineBoxes;
          markerUtf16Len = cached.markerUtf16Len;
        } else {
          if (cached?.paragraph?.delete) {
            try {
              cached.paragraph.delete();
            } catch {
              // Ignore stale CanvasKit paragraph cleanup errors.
            }
          }
          para = this.buildParagraph(item, model, columnWidth);
          markerUtf16Len = item.numbering?.text.length ?? 0;
          lineBoxes = this.extractParagraphLineBoxes(para, markerUtf16Len);
          this.paragraphMeasureCache.set(i, {
            key: cacheKey,
            paragraph: para,
            lineBoxes,
            markerUtf16Len,
          });
        }

        results.push({
          bodyIndex: i,
          item,
          section,
          sectionIndex: sectionIdx,
          paragraph: para,
          lineBoxes,
          markerUtf16Len,
        });
      } else {
        seenTableIndices.add(i);
        const cacheKey = this.tableMeasureCacheKey(item, section, columnWidth);
        const cached = this.tableMeasureCache.get(i);
        const tableInfo =
          cached && cached.key === cacheKey
            ? cached.tableInfo
            : this.measureTable(item);
        if (!cached || cached.key !== cacheKey) {
          this.tableMeasureCache.set(i, { key: cacheKey, tableInfo });
        }
        results.push({
          bodyIndex: i,
          item,
          section,
          sectionIndex: sectionIdx,
          tableInfo,
        });
      }
    }

    this.pruneMeasurementCaches(seenParagraphIndices, seenTableIndices);
    return results;
  }

  private paragraphMeasureCacheKey(
    paragraph: ParagraphModel,
    section: SectionModel,
    columnWidthPt: number,
  ): string {
    return JSON.stringify({
      sectionIndex: paragraph.sectionIndex,
      sectionPageWidth: section.pageWidthPt,
      sectionPageHeight: section.pageHeightPt,
      sectionMargins: section.margins,
      columnCount: section.columnCount ?? 1,
      columnWidthPt,
      paragraph,
    });
  }

  private tableMeasureCacheKey(
    table: TableModel,
    section: SectionModel,
    columnWidthPt: number,
  ): string {
    return JSON.stringify({
      sectionIndex: table.sectionIndex,
      sectionPageWidth: section.pageWidthPt,
      sectionPageHeight: section.pageHeightPt,
      sectionMargins: section.margins,
      columnCount: section.columnCount ?? 1,
      columnWidthPt,
      table,
    });
  }

  private pruneMeasurementCaches(
    seenParagraphIndices: Set<number>,
    seenTableIndices: Set<number>,
  ): void {
    for (const [index, entry] of this.paragraphMeasureCache) {
      if (seenParagraphIndices.has(index)) continue;
      if (entry.paragraph?.delete) {
        try {
          entry.paragraph.delete();
        } catch {
          // Ignore cleanup failures.
        }
      }
      this.paragraphMeasureCache.delete(index);
    }

    for (const index of this.tableMeasureCache.keys()) {
      if (seenTableIndices.has(index)) continue;
      this.tableMeasureCache.delete(index);
    }
  }

  /**
   * Build a CanvasKit Paragraph from a ParagraphModel.
   */
  private buildParagraph(
    p: ParagraphModel,
    _model: DocViewModel,
    contentWidthPt: number,
  ): any {
    const ck = this.ck;

    const effectiveWidth = this.paragraphEffectiveWidth(p, contentWidthPt);
    const paraStyleInput: Record<string, unknown> = {
      textAlign: this.mapAlignment(p.alignment),
      textDirection: ck.TextDirection.LTR,
      textStyle: { color: ck.Color(0, 0, 0, 1) },
      replaceTabCharacters: false,
    };
    if (p.lineSpacing?.rule === "auto" && p.lineSpacing.value > 0) {
      paraStyleInput.heightMultiplier = p.lineSpacing.value;
    } else if (
      p.lineSpacing &&
      (p.lineSpacing.rule === "exact" || p.lineSpacing.rule === "atLeast") &&
      p.lineSpacing.value > 0
    ) {
      const base = this.getEffectiveFontSize(p);
      const multiplier = Math.max(0.6, p.lineSpacing.value / Math.max(base, 1));
      paraStyleInput.heightMultiplier = multiplier;
    }
    const paraStyle = new ck.ParagraphStyle(paraStyleInput);
    const tokens = this.buildParagraphTokens(p);
    const tabCount = tokens.filter(
      (t) => t.kind === "placeholder" && t.role === "tab",
    ).length;
    const { defaultTabStopPt, explicitStopsPt } = this.resolveParagraphTabStops(
      p,
      effectiveWidth,
    );
    let tabWidthsPt = new Array<number>(tabCount).fill(defaultTabStopPt);
    let paragraph = this.layoutParagraphFromTokens(
      paraStyle,
      tokens,
      effectiveWidth,
      tabWidthsPt,
    );

    if (tabCount === 0) {
      return paragraph;
    }

    for (let pass = 0; pass < MAX_LAYOUT_REFINEMENT_PASSES; pass++) {
      const nextTabWidths = this.refineTabWidthsFromLayout(
        paragraph,
        tokens,
        tabWidthsPt,
        defaultTabStopPt,
        explicitStopsPt,
      );
      if (!nextTabWidths || this.sameNumberArray(nextTabWidths, tabWidthsPt)) {
        return paragraph;
      }
      tabWidthsPt = nextTabWidths;
      paragraph.delete();
      paragraph = this.layoutParagraphFromTokens(
        paraStyle,
        tokens,
        effectiveWidth,
        tabWidthsPt,
      );
    }
    return paragraph;
  }

  private paragraphEffectiveWidth(
    p: ParagraphModel,
    contentWidthPt: number,
  ): number {
    let effectiveWidth = contentWidthPt;
    if (p.indents) {
      if (p.indents.leftPt) effectiveWidth -= p.indents.leftPt;
      if (p.indents.rightPt) effectiveWidth -= p.indents.rightPt;
    }
    effectiveWidth -= this.floatingWrapReductionPt(p, contentWidthPt);
    return Math.max(50, effectiveWidth);
  }

  private floatingWrapReductionPt(
    p: ParagraphModel,
    contentWidthPt: number,
  ): number {
    let maxReduction = 0;
    for (const run of p.runs ?? []) {
      const floating = run.floatingImage;
      if (!floating) continue;
      const wrapType = (floating.wrapType ?? "square").toLowerCase();
      if (
        wrapType.includes("none") ||
        wrapType.includes("behind") ||
        wrapType.includes("infront")
      ) {
        continue;
      }
      maxReduction = Math.max(maxReduction, floating.widthPt);
    }

    if (maxReduction <= 0) return 0;
    // Keep at least 40% of line width for text flow in v1.
    return Math.min(maxReduction, contentWidthPt * 0.6);
  }

  private buildParagraphTokens(p: ParagraphModel): ParagraphToken[] {
    const ck = this.ck;
    const tokens: ParagraphToken[] = [];
    let tabIndex = 0;

    if (p.numbering) {
      const style = this.makeTextStyle({} as RunModel, p);
      tokens.push({
        kind: "text",
        text: p.numbering.text,
        style,
        utf16Len: p.numbering.text.length,
      });
    }

    for (const run of p.runs ?? []) {
      if (run.inlineImage) {
        tokens.push({
          kind: "placeholder",
          role: "inlineImage",
          widthPt: run.inlineImage.widthPt,
          heightPt: run.inlineImage.heightPt,
          alignment: ck.PlaceholderAlignment.Bottom,
          baseline: ck.TextBaseline.Alphabetic,
          baselineOffsetPt: 0,
        });
        continue;
      }

      if (run.hasBreak) {
        if (run.text) {
          const style = this.makeTextStyle(run, p);
          tokens.push({
            kind: "text",
            text: run.text,
            style,
            utf16Len: run.text.length,
          });
        }
        const style = this.makeTextStyle(run, p);
        tokens.push({
          kind: "text",
          text: "\n",
          style,
          utf16Len: 1,
        });
        continue;
      }

      if (run.hasTab) {
        if (run.text) {
          const style = this.makeTextStyle(run, p);
          tokens.push({
            kind: "text",
            text: run.text,
            style,
            utf16Len: run.text.length,
          });
        }
        const fontSizePt = Math.max(
          1,
          run.fontSizePt ?? this.getEffectiveFontSize(p),
        );
        tokens.push({
          kind: "placeholder",
          role: "tab",
          widthPt: DEFAULT_TAB_STOP_PT,
          heightPt: fontSizePt * 1.2,
          alignment: ck.PlaceholderAlignment.Baseline,
          baseline: ck.TextBaseline.Alphabetic,
          baselineOffsetPt: 0,
          tabIndex,
        });
        tabIndex += 1;
        continue;
      }

      if (run.footnoteRef != null || run.endnoteRef != null) {
        const refText = String(run.footnoteRef ?? run.endnoteRef ?? "");
        const paragraphDefaults = this.headingTextDefaults(p) ?? {};
        const style = this.makeTextStyle(
          {
            ...paragraphDefaults,
            fontSizePt:
              (run.fontSizePt ??
                paragraphDefaults.fontSizePt ??
                DEFAULT_FONT_SIZE_PT) * 0.75,
            superscript: true,
            color: "0563C1",
          } as RunModel,
          p,
        );
        tokens.push({
          kind: "text",
          text: refText,
          style,
          utf16Len: refText.length,
        });
        continue;
      }

      // Normal text run
      const text = run.text || "";
      if (text.length > 0) {
        const style = this.makeTextStyle(run, p);
        tokens.push({
          kind: "text",
          text,
          style,
          utf16Len: text.length,
        });
      }
    }

    return tokens;
  }

  private layoutParagraphFromTokens(
    paraStyle: any,
    tokens: ParagraphToken[],
    effectiveWidth: number,
    tabWidthsPt: number[],
  ): any {
    const ck = this.ck;
    const builder = ck.ParagraphBuilder.MakeFromFontCollection(
      paraStyle,
      this.fontCollection,
    );

    for (const token of tokens) {
      if (token.kind === "text") {
        builder.pushStyle(token.style);
        builder.addText(token.text);
        builder.pop();
        continue;
      }

      if (token.role === "tab") {
        const tabIndex = token.tabIndex ?? 0;
        const widthPt = Math.max(
          1,
          tabWidthsPt[tabIndex] ?? DEFAULT_TAB_STOP_PT,
        );
        builder.addPlaceholder(
          widthPt,
          token.heightPt,
          token.alignment,
          token.baseline,
          token.baselineOffsetPt,
        );
        continue;
      }

      builder.addPlaceholder(
        token.widthPt,
        token.heightPt,
        token.alignment,
        token.baseline,
        token.baselineOffsetPt,
      );
    }

    const paragraph = builder.build();
    paragraph.layout(effectiveWidth);
    return paragraph;
  }

  private resolveParagraphTabStops(
    p: ParagraphModel,
    paragraphWidthPt: number,
  ): { defaultTabStopPt: number; explicitStopsPt: number[] } {
    const defaultTabStopPt =
      p.defaultTabStopPt && p.defaultTabStopPt > 0
        ? p.defaultTabStopPt
        : DEFAULT_TAB_STOP_PT;
    const rawStops = p.tabStops ?? [];
    const stops: number[] = [];
    for (const stop of rawStops) {
      const pos =
        typeof stop === "number"
          ? stop
          : (stop.positionPt ?? stop.posPt ?? stop.valuePt);
      if (!Number.isFinite(pos) || pos == null || pos <= 0) continue;
      if (pos > paragraphWidthPt + defaultTabStopPt) continue;
      stops.push(pos);
    }
    stops.sort((a, b) => a - b);
    const deduped: number[] = [];
    for (const stop of stops) {
      const prev = deduped[deduped.length - 1];
      if (prev == null || Math.abs(stop - prev) > 1e-3) {
        deduped.push(stop);
      }
    }
    return { defaultTabStopPt, explicitStopsPt: deduped };
  }

  private refineTabWidthsFromLayout(
    paragraph: any,
    tokens: ParagraphToken[],
    previousTabWidthsPt: number[],
    defaultTabStopPt: number,
    explicitStopsPt: number[],
  ): number[] | null {
    const metrics = paragraph.getLineMetrics?.() as
      | Array<{
          startIndex: number;
          endIndex: number;
          endIncludingNewline: number;
          left: number;
        }>
      | undefined;
    const placeholders = paragraph.getRectsForPlaceholders?.() as
      | Array<{ rect: [number, number, number, number] }>
      | undefined;
    if (!metrics || !placeholders || placeholders.length === 0) {
      return null;
    }

    const nextWidths = previousTabWidthsPt.slice();
    let utf16Cursor = 0;
    let placeholderCursor = 0;
    for (const token of tokens) {
      if (token.kind === "text") {
        utf16Cursor += token.utf16Len;
        continue;
      }

      const placeholderRect = placeholders[placeholderCursor];
      const rect = placeholderRect?.rect;
      const tokenOffset = utf16Cursor;
      utf16Cursor += 1;
      placeholderCursor += 1;

      if (token.role !== "tab") continue;
      if (!rect) continue;

      const line = metrics.find((m) => {
        const end = Math.max(m.endIncludingNewline, m.endIndex);
        return tokenOffset >= m.startIndex && tokenOffset < end;
      });
      const lineLeft = line?.left ?? 0;
      const xPt = Math.max(0, rect[0] - lineLeft);
      const nextStop = this.nextTabStopPt(
        xPt,
        explicitStopsPt,
        defaultTabStopPt,
      );
      const tabIndex = token.tabIndex ?? 0;
      nextWidths[tabIndex] = Math.max(1, nextStop - xPt);
    }

    return nextWidths;
  }

  private nextTabStopPt(
    xPt: number,
    explicitStopsPt: number[],
    defaultTabStopPt: number,
  ): number {
    const epsilon = 1e-4;
    for (const stop of explicitStopsPt) {
      if (stop > xPt + epsilon) return stop;
    }
    const step = Math.max(1, defaultTabStopPt);
    const bucket = Math.floor((xPt + epsilon) / step) + 1;
    return bucket * step;
  }

  /**
   * Build a CanvasKit TextStyle from a RunModel.
   */
  private makeTextStyle(run: RunModel, paragraph?: ParagraphModel): any {
    const ck = this.ck;
    const paragraphDefaults = paragraph
      ? this.headingTextDefaults(paragraph)
      : null;
    const family = resolveFontFamily(
      run.fontFamily ?? paragraphDefaults?.fontFamily,
    );
    const fontSize =
      run.fontSizePt ?? paragraphDefaults?.fontSizePt ?? DEFAULT_FONT_SIZE_PT;
    const hyperlink =
      typeof run.hyperlink === "string" && run.hyperlink.length > 0;

    let color = ck.Color(0, 0, 0, 1);
    const colorHex =
      run.color ??
      paragraphDefaults?.color ??
      (hyperlink ? "0563C1" : undefined);
    if (colorHex) {
      color = this.parseHexColor(colorHex);
    }

    const decoration: number =
      ((run.underline ?? hyperlink) ? ck.UnderlineDecoration : 0) |
      (run.strikethrough ? ck.LineThroughDecoration : 0);

    const style: any = {
      fontFamilies: [family],
      fontSize,
      color,
      fontStyle: {
        weight:
          (run.bold ?? paragraphDefaults?.bold)
            ? ck.FontWeight.Bold
            : ck.FontWeight.Normal,
        slant:
          (run.italic ?? paragraphDefaults?.italic)
            ? ck.FontSlant.Italic
            : ck.FontSlant.Upright,
      },
    };

    if (decoration) {
      style.decoration = decoration;
      style.decorationColor = color;
    }

    if (run.highlight) {
      style.backgroundColor = this.highlightColor(run.highlight);
    }

    return new ck.TextStyle(style);
  }

  /**
   * Compute an approximate table height from its row definitions.
   */
  private measureTable(t: TableModel): {
    rows: Array<{
      heightPt: number;
      isHeader?: boolean;
      repeatHeader?: boolean;
      keepTogether?: boolean;
      keepWithNext?: boolean;
      cantSplit?: boolean;
    }>;
    totalHeightPt: number;
  } {
    let total = 0;
    const rows: Array<{
      heightPt: number;
      isHeader?: boolean;
      repeatHeader?: boolean;
      keepTogether?: boolean;
      keepWithNext?: boolean;
      cantSplit?: boolean;
    }> = [];
    for (const row of t.rows) {
      const h = row.heightPt ?? TABLE_DEFAULT_ROW_HEIGHT_PT;
      rows.push({
        heightPt: h,
        isHeader: row.isHeader,
        repeatHeader: row.repeatHeader,
        keepTogether: row.keepTogether,
        keepWithNext: row.keepWithNext,
        cantSplit: row.cantSplit,
      });
      total += h;
    }
    return { rows, totalHeightPt: total };
  }

  private sectionColumnWidth(section: SectionModel): number {
    const contentWidth =
      section.pageWidthPt - section.margins.left - section.margins.right;
    const columns = Math.max(1, Math.floor(section.columnCount ?? 1));
    const gapPt = 18;
    const totalGap = gapPt * Math.max(0, columns - 1);
    return Math.max(32, (contentWidth - totalGap) / columns);
  }

  private extractParagraphLineBoxes(
    paragraph: any,
    markerUtf16Len = 0,
  ): ParagraphLineBox[] {
    const metrics = paragraph.getLineMetrics?.() as
      | Array<{
          lineNumber: number;
          startIndex: number;
          endIndex: number;
          endIncludingNewline: number;
          baseline: number;
          ascent: number;
          descent: number;
          left: number;
          width: number;
        }>
      | undefined;
    if (!metrics || metrics.length === 0) {
      const h = paragraph.getHeight?.() ?? 14;
      return [
        {
          lineIndex: 0,
          startUtf16: 0,
          endUtf16: 0,
          topPt: 0,
          bottomPt: Math.max(1, h),
          baselinePt: Math.max(1, h) * 0.8,
          leftPt: 0,
          widthPt: paragraph.getLongestLine?.() ?? 0,
        },
      ];
    }

    const lines: ParagraphLineBox[] = [];
    for (const m of metrics) {
      const top = m.baseline - Math.abs(m.ascent);
      const bottom = m.baseline + Math.abs(m.descent);
      lines.push({
        lineIndex: m.lineNumber,
        startUtf16: Math.max(0, m.startIndex - markerUtf16Len),
        endUtf16: Math.max(
          0,
          Math.max(m.endIndex, m.endIncludingNewline) - markerUtf16Len,
        ),
        topPt: top,
        bottomPt: bottom,
        baselinePt: m.baseline,
        leftPt: m.left,
        widthPt: m.width,
      });
    }
    return lines;
  }

  /**
   * Get the effective font size from heading level or default.
   */
  private getEffectiveFontSize(p: ParagraphModel): number {
    if (p.headingLevel) {
      switch (p.headingLevel) {
        case 1:
          return 16;
        case 2:
          return 13;
        case 3:
          return 12;
        case 4:
          return 11;
        case 5:
          return 11;
        case 6:
          return 10.5;
      }
    }
    return DEFAULT_FONT_SIZE_PT;
  }

  private headingTextDefaults(
    p: ParagraphModel,
  ): Pick<
    RunModel,
    "bold" | "italic" | "fontFamily" | "fontSizePt" | "color"
  > | null {
    switch (p.headingLevel) {
      case 1:
        return { bold: true, fontSizePt: 16, color: "2F5496" };
      case 2:
        return { bold: true, fontSizePt: 13, color: "2F5496" };
      case 3:
        return { bold: true, fontSizePt: 12, color: "1F3763" };
      case 4:
        return { bold: true, italic: true, fontSizePt: 11, color: "2F5496" };
      case 5:
        return { fontSizePt: 11, color: "2F5496" };
      case 6:
        return { italic: true, fontSizePt: 10.5, color: "1F3763" };
      default:
        return null;
    }
  }

  // ==========================================================================
  // Internal: drawing
  // ==========================================================================

  /** Draw all currently visible pages. */
  private drawVisiblePages(): void {
    this.pool.syncDevicePixelRatio();
    const range = this.pool.getVisibleRange();
    for (let i = range.start; i <= range.end; i++) {
      this.drawPage(i);
    }
  }

  /** Draw a single page onto its canvas. */
  private drawPage(pageIndex: number): void {
    const canvas = this.pool.getCanvas(pageIndex);
    if (!canvas) return;

    const page = this.pages[pageIndex];
    if (!page) return;

    const ck = this.ck;
    const surface =
      ck.MakeWebGLCanvasSurface(canvas) ?? ck.MakeSWCanvasSurface(canvas);
    if (!surface) return;

    const skCanvas = surface.getCanvas();
    const section = page.section;

    // Scale so the drawing coordinate system is in POINTS.
    // The canvas CSS size is pageWidthPt "pt", which the browser renders as
    // pageWidthPt * 96/72 CSS px. The buffer is that * current device DPR.
    // By scaling by (buffer / pageWidthPt) we get 1 drawing unit = 1 pt,
    // so fontSize: 11 renders as 11pt visually.
    const scaleX = canvas.width / section.pageWidthPt;
    const scaleY = canvas.height / section.pageHeightPt;
    skCanvas.scale(scaleX, scaleY);

    // White background
    const white = ck.Color(255, 255, 255, 1);
    skCanvas.clear(white);

    // Content area origin in points (no conversion needed)
    const originX = section.margins.left;
    const originY = section.margins.top;

    // Draw page items (all coordinates are in points)
    for (const pi of page.items) {
      if (pi.type === "paragraph") {
        const lp = pi.ref as LaidOutParagraph;
        const px = originX + lp.xPt;
        const py = originY + lp.yPt;
        const ckRect = ck.XYWHRect(
          px - PARAGRAPH_CLIP_PAD_PT,
          py - PARAGRAPH_CLIP_PAD_PT,
          lp.widthPt + PARAGRAPH_CLIP_PAD_PT * 2,
          lp.heightPt + PARAGRAPH_CLIP_PAD_PT * 2,
        );
        const sourceY = py - lp.sourceTopPt;
        skCanvas.save();
        skCanvas.clipRect(ckRect, ck.ClipOp.Intersect, true);
        skCanvas.drawParagraph(lp.paragraph, px, sourceY);
        this.drawInlineImages(skCanvas, lp, px, sourceY, pageIndex);
        skCanvas.restore();
      } else {
        const lt = pi.ref as LaidOutTable;
        this.drawTable(skCanvas, lt, originX, originY, pageIndex);
      }
    }

    // Floating images are paragraph-anchored but drawn as independent objects.
    for (const pi of page.items) {
      if (pi.type !== "paragraph") continue;
      this.drawFloatingImages(
        skCanvas,
        pi.ref as LaidOutParagraph,
        originX,
        originY,
      );
    }

    if (this.selection) {
      this.drawSelectionHighlights(skCanvas, pageIndex, originX, originY);
    }

    surface.flush();
    surface.delete();
  }

  /**
   * Draw inline images for placeholders within a paragraph.
   */
  private drawInlineImages(
    skCanvas: any,
    lp: LaidOutParagraph,
    paraXPt: number,
    paraYPt: number,
    _pageIndex: number,
  ): void {
    if (!this.currentModel) return;

    const bodyItem = this.currentModel.body[lp.bodyIndex];
    if (!bodyItem || bodyItem.type !== "paragraph") return;

    const placeholders = lp.paragraph.getRectsForPlaceholders?.();
    if (!placeholders || placeholders.length === 0) return;

    let placeholderIdx = 0;
    for (const run of bodyItem.runs) {
      if (run.hasTab && !run.inlineImage) {
        placeholderIdx += 1;
      }
      if (!run.inlineImage) continue;
      if (placeholderIdx >= placeholders.length) break;

      const rect = placeholders[placeholderIdx];
      placeholderIdx++;
      if (!rect?.rect) continue;

      const imgData = this.currentModel.images[run.inlineImage.imageIndex];
      if (!imgData) continue;

      const skImage = this.getOrDecodeImage(
        run.inlineImage.imageIndex,
        imgData.dataUri,
      );
      if (!skImage) continue;

      // All coordinates are in points (paragraph layout + drawing coord system)
      const destX = paraXPt + rect.rect[0];
      const destY = paraYPt + rect.rect[1];
      const destW = run.inlineImage.widthPt;
      const destH = run.inlineImage.heightPt;

      const ck = this.ck;
      const srcRect = ck.XYWHRect(0, 0, skImage.width(), skImage.height());
      const dstRect = ck.XYWHRect(destX, destY, destW, destH);
      const paint = new ck.Paint();
      skCanvas.drawImageRect(skImage, srcRect, dstRect, paint, false);
      paint.delete();
    }
  }

  /**
   * Draw paragraph-anchored floating images.
   * v1 behavior: deterministic placement by paragraph origin + offsets.
   */
  private drawFloatingImages(
    skCanvas: any,
    lp: LaidOutParagraph,
    originXPt: number,
    originYPt: number,
  ): void {
    if (!this.currentModel) return;
    const bodyItem = this.currentModel.body[lp.bodyIndex];
    if (!bodyItem || bodyItem.type !== "paragraph") return;

    for (const run of bodyItem.runs) {
      const floating = run.floatingImage;
      if (!floating) continue;

      const imgData = this.currentModel.images[floating.imageIndex];
      if (!imgData) continue;
      const skImage = this.getOrDecodeImage(
        floating.imageIndex,
        imgData.dataUri,
      );
      if (!skImage) continue;

      const destX = originXPt + lp.xPt + floating.offsetXPt;
      const destY = originYPt + lp.yPt + floating.offsetYPt - lp.sourceTopPt;
      const ck = this.ck;
      const srcRect = ck.XYWHRect(0, 0, skImage.width(), skImage.height());
      const dstRect = ck.XYWHRect(
        destX,
        destY,
        Math.max(1, floating.widthPt),
        Math.max(1, floating.heightPt),
      );
      const paint = new ck.Paint();
      skCanvas.drawImageRect(skImage, srcRect, dstRect, paint, false);
      paint.delete();
    }
  }

  /**
   * Draw a table: grid lines and cell text.
   */
  private drawTable(
    skCanvas: any,
    lt: LaidOutTable,
    originXPt: number,
    originYPt: number,
    _pageIndex: number,
  ): void {
    if (!this.currentModel) return;
    const bodyItem = this.currentModel.body[lt.bodyIndex];
    if (!bodyItem || bodyItem.type !== "table") return;

    const ck = this.ck;
    const table = bodyItem;
    const tableX = originXPt + lt.xPt;
    let tableY = originYPt + lt.yPt;

    // Use the same column width fallback as hit-testing/rect lookup so
    // bare editor-authored tables still render visible columns.
    const colWidths = this.tableColumnWidthsPt(table, lt);
    const colXPts: number[] = [0];
    let accum = 0;
    for (const w of colWidths) {
      accum += w;
      colXPts.push(accum);
    }
    const totalWidthPt = accum || lt.widthPt;

    // Border paint
    const borderPaint = new ck.Paint();
    borderPaint.setColor(ck.Color(0, 0, 0, 1));
    borderPaint.setStyle(ck.PaintStyle.Stroke);
    borderPaint.setStrokeWidth(0.75);
    borderPaint.setAntiAlias(true);

    // Cell text paint
    const textPaint = new ck.Paint();
    textPaint.setColor(ck.Color(0, 0, 0, 1));
    textPaint.setAntiAlias(true);

    const rows: TableRowModel[] = [];
    if ((lt.repeatedHeaderRowCount ?? 0) > 0) {
      rows.push(...table.rows.slice(0, lt.repeatedHeaderRowCount));
    }
    rows.push(...table.rows.slice(lt.rowStart, lt.rowEnd));
    for (const row of rows) {
      const rowH = row.heightPt ?? TABLE_DEFAULT_ROW_HEIGHT_PT;

      let colIdx = 0;
      for (const cell of row.cells) {
        if (cell.isCovered) {
          colIdx += cell.colSpan ?? 1;
          continue;
        }

        const colSpan = cell.colSpan ?? 1;
        const cellXPt = colXPts[colIdx] ?? 0;
        const cellEndXPt = colXPts[colIdx + colSpan] ?? totalWidthPt;
        const cellW = cellEndXPt - cellXPt;

        const cellX = tableX + cellXPt;

        // Draw cell background
        if (cell.shadingColor) {
          const bgPaint = new ck.Paint();
          bgPaint.setColor(this.parseHexColor(cell.shadingColor));
          bgPaint.setStyle(ck.PaintStyle.Fill);
          skCanvas.drawRect(ck.XYWHRect(cellX, tableY, cellW, rowH), bgPaint);
          bgPaint.delete();
        }

        // Draw cell border
        skCanvas.drawRect(ck.XYWHRect(cellX, tableY, cellW, rowH), borderPaint);

        // Draw cell text (all in points)
        if (cell.text) {
          const paraStyle = new ck.ParagraphStyle({
            textAlign: ck.TextAlign.Left,
            textStyle: { color: ck.Color(0, 0, 0, 1) },
          });
          const builder = ck.ParagraphBuilder.MakeFromFontCollection(
            paraStyle,
            this.fontCollection,
          );
          const textStyle = new ck.TextStyle({
            fontFamilies: [resolveFontFamily()],
            fontSize: DEFAULT_FONT_SIZE_PT,
            color: ck.Color(0, 0, 0, 1),
          });
          builder.pushStyle(textStyle);
          builder.addText(cell.text);
          builder.pop();

          const cellPara = builder.build();
          const layoutWidth = Math.max(cellW - TABLE_CELL_PADDING_PT * 2, 10);
          cellPara.layout(layoutWidth);

          skCanvas.drawParagraph(
            cellPara,
            cellX + TABLE_CELL_PADDING_PT,
            tableY + TABLE_CELL_PADDING_PT,
          );
          cellPara.delete();
        }

        colIdx += colSpan;
      }

      tableY += rowH;
    }

    borderPaint.delete();
    textPaint.delete();
  }

  // ==========================================================================
  // Internal: hit testing
  // ==========================================================================

  private hitTestPage(
    pageIndex: number,
    localXPt: number,
    localYPt: number,
  ): HitTestResult | null {
    const page = this.pages[pageIndex];
    if (!page) return null;

    const section = page.section;

    // Convert to content-area-relative coordinates
    const contentXPt = localXPt - section.margins.left;
    const contentYPt = localYPt - section.margins.top;

    const paragraphs = page.items
      .filter((pi): pi is { type: "paragraph"; ref: LaidOutParagraph } => {
        return pi.type === "paragraph";
      })
      .map((pi) => pi.ref)
      .sort((a, b) => a.yPt - b.yPt);

    if (paragraphs.length === 0) return null;

    // If click lands on a paragraph's vertical band, place within it.
    // This includes clicks in left/right margins, matching Docs/Word.
    for (const lp of paragraphs) {
      if (contentYPt >= lp.yPt && contentYPt <= lp.yPt + lp.heightPt) {
        return this.hitTestParagraph(lp, contentXPt, contentYPt);
      }
    }

    // Click above first paragraph on page -> paragraph start.
    const first = paragraphs[0]!;
    if (contentYPt < first.yPt) {
      return {
        bodyIndex: first.bodyIndex,
        charOffset: first.utf16Start,
        affinity: "leading",
      };
    }

    // Click below last paragraph on page -> paragraph end.
    const last = paragraphs[paragraphs.length - 1]!;
    const lastBottom = last.yPt + last.heightPt;
    if (contentYPt > lastBottom) {
      return {
        bodyIndex: last.bodyIndex,
        charOffset: this.paragraphEndOffset(last),
        affinity: "trailing",
      };
    }

    // Click in gap between paragraphs -> nearest boundary around gap center.
    for (let i = 1; i < paragraphs.length; i++) {
      const prev = paragraphs[i - 1]!;
      const curr = paragraphs[i]!;
      const prevBottom = prev.yPt + prev.heightPt;

      if (contentYPt >= prevBottom && contentYPt < curr.yPt) {
        const mid = (prevBottom + curr.yPt) / 2;
        if (contentYPt < mid) {
          return {
            bodyIndex: prev.bodyIndex,
            charOffset: this.paragraphEndOffset(prev),
            affinity: "trailing",
          };
        }
        return {
          bodyIndex: curr.bodyIndex,
          charOffset: curr.utf16Start,
          affinity: "leading",
        };
      }
    }

    return null;
  }

  private hitTestTableCell(
    pageIndex: number,
    localXPt: number,
    localYPt: number,
  ): TableCellPosition | null {
    if (!this.currentModel) return null;
    const page = this.pages[pageIndex];
    if (!page) return null;
    const contentXPt = localXPt - page.section.margins.left;
    const contentYPt = localYPt - page.section.margins.top;
    for (const lt of this.laidOutTables) {
      if (lt.pageIndex !== pageIndex) continue;
      const bodyItem = this.currentModel.body[lt.bodyIndex];
      if (!bodyItem || bodyItem.type !== "table") continue;
      const widthPt = lt.widthPt;
      if (
        contentXPt < lt.xPt ||
        contentXPt > lt.xPt + widthPt ||
        contentYPt < lt.yPt ||
        contentYPt > lt.yPt + lt.heightPt
      ) {
        continue;
      }
      const colWidths = this.tableColumnWidthsPt(bodyItem, lt);
      let rowY = lt.yPt;
      for (let row = lt.rowStart; row < lt.rowEnd; row += 1) {
        const rowModel = bodyItem.rows[row];
        const rowHeight = rowModel?.heightPt ?? TABLE_DEFAULT_ROW_HEIGHT_PT;
        if (contentYPt >= rowY && contentYPt <= rowY + rowHeight) {
          let colX = lt.xPt;
          for (let col = 0; col < colWidths.length; col += 1) {
            const colWidth = colWidths[col] ?? 0;
            if (contentXPt >= colX && contentXPt <= colX + colWidth) {
              return { bodyIndex: lt.bodyIndex, row, col };
            }
            colX += colWidth;
          }
          return null;
        }
        rowY += rowHeight;
      }
    }
    return null;
  }

  private tableCellRectInPagePt(
    table: Extract<BodyItem, { type: "table" }>,
    lt: LaidOutTable,
    page: PageInfo,
    cell: TableCellPosition,
  ): { xPt: number; yPt: number; wPt: number; hPt: number } | null {
    const rowModel = table.rows[cell.row];
    if (!rowModel) return null;
    const colWidths = this.tableColumnWidthsPt(table, lt);
    if (cell.col < 0 || cell.col >= colWidths.length) return null;
    let xOffset = 0;
    for (let col = 0; col < cell.col; col += 1) {
      xOffset += colWidths[col] ?? 0;
    }
    let yOffset = 0;
    for (let row = lt.rowStart; row < cell.row; row += 1) {
      yOffset += table.rows[row]?.heightPt ?? TABLE_DEFAULT_ROW_HEIGHT_PT;
    }
    return {
      xPt: page.section.margins.left + lt.xPt + xOffset,
      yPt: page.section.margins.top + lt.yPt + yOffset,
      wPt: colWidths[cell.col] ?? 0,
      hPt: rowModel.heightPt ?? TABLE_DEFAULT_ROW_HEIGHT_PT,
    };
  }

  private tableColumnWidthsPt(
    table: Extract<BodyItem, { type: "table" }>,
    lt: LaidOutTable,
  ): number[] {
    if (table.columnWidthsPt.length > 0) {
      return table.columnWidthsPt.slice();
    }
    const columns = table.rows[0]?.cells.length ?? 0;
    if (columns <= 0) return [];
    return Array.from({ length: columns }, () => lt.widthPt / columns);
  }

  private hitTestParagraph(
    lp: LaidOutParagraph,
    contentXPt: number,
    contentYPt: number,
  ): HitTestResult {
    const relX = contentXPt - lp.xPt;
    const relY = contentYPt - lp.yPt + lp.sourceTopPt;
    const pos = lp.paragraph.getGlyphPositionAtCoordinate(relX, relY);
    const min = lp.utf16Start;
    const max = Math.max(lp.utf16End, lp.utf16Start);
    const minLayout = min + lp.markerUtf16Len;
    const maxLayout = max + lp.markerUtf16Len;
    const clampedLayoutOffset = Math.max(
      minLayout,
      Math.min(pos?.pos ?? minLayout, maxLayout),
    );
    const clampedOffset = Math.max(
      min,
      Math.min(clampedLayoutOffset - lp.markerUtf16Len, max),
    );

    return {
      bodyIndex: lp.bodyIndex,
      charOffset: clampedOffset,
      affinity:
        clampedOffset >= max
          ? "trailing"
          : pos?.affinity === 1
            ? "trailing"
            : "leading",
    };
  }

  private paragraphEndOffset(lp: LaidOutParagraph): number {
    return Math.max(lp.utf16End, lp.utf16Start);
  }

  // ==========================================================================
  // Internal: cursor and selection
  // ==========================================================================

  /**
   * Draw selection highlight rectangles on the given page.
   * All coordinates are in points.
   */
  private drawSelectionHighlights(
    skCanvas: any,
    pageIndex: number,
    originX: number,
    originY: number,
  ): void {
    if (!this.selection) return;

    const ck = this.ck;
    const normalized = normalizeSelectionRange(this.selection);
    if (!normalized) return;

    const paint = new ck.Paint();
    paint.setColor(
      ck.Color(
        SELECTION_COLOR_RGBA[0],
        SELECTION_COLOR_RGBA[1],
        SELECTION_COLOR_RGBA[2],
        SELECTION_COLOR_RGBA[3],
      ),
    );
    paint.setStyle(ck.PaintStyle.Fill);
    paint.setAntiAlias(true);

    for (const lp of this.laidOutParagraphs) {
      if (lp.pageIndex !== pageIndex) continue;
      const paragraphEnd = this.getParagraphCharCountByBodyIndex(lp.bodyIndex);
      const range = selectionRangeInParagraphFragment(
        normalized,
        lp.bodyIndex,
        lp.utf16Start,
        lp.utf16End,
        paragraphEnd,
      );
      if (!range) continue;

      const rects = lp.paragraph.getRectsForRange(
        range.start + lp.markerUtf16Len,
        range.end + lp.markerUtf16Len,
        ck.RectHeightStyle.Max,
        ck.RectWidthStyle.Tight,
      );
      for (const rect of extractRectList(rects as unknown)) {
        const rx = originX + lp.xPt + rect[0];
        const ry = originY + lp.yPt + rect[1] - lp.sourceTopPt;
        const rw = rect[2] - rect[0];
        const rh = rect[3] - rect[1];
        if (!(rw > 0 && rh > 0)) continue;
        skCanvas.drawRect(ck.XYWHRect(rx, ry, rw, rh), paint);
      }
    }

    paint.delete();
  }

  private findLaidOutParagraphForPosition(
    pos: DocPosition,
  ): LaidOutParagraph | null {
    const charOffset = Math.max(0, pos.charOffset);
    const matches = this.paragraphsByBodyIndex.get(pos.bodyIndex) ?? [];

    if (matches.length === 0) {
      return this.findNearestParagraphByBodyIndex(pos.bodyIndex);
    }
    if (matches.length === 1) return matches[0] ?? null;

    const para = matches[0]?.paragraph;
    const markerUtf16Len = matches[0]?.markerUtf16Len ?? 0;
    const layoutOffset = charOffset + markerUtf16Len;
    let lineNo = -1;
    if (para?.getLineNumberAt) {
      lineNo = para.getLineNumberAt(layoutOffset);
      if (lineNo < 0 && layoutOffset > 0) {
        lineNo = para.getLineNumberAt(layoutOffset - 1);
      }
    }

    if (lineNo >= 0) {
      for (const lp of matches) {
        if (lineNo >= lp.lineStart && lineNo < lp.lineEnd) {
          return lp;
        }
      }
    }

    let nearest: LaidOutParagraph | null = null;
    let nearestDistance = Number.POSITIVE_INFINITY;
    for (let i = 0; i < matches.length; i++) {
      const lp = matches[i]!;
      const end = Math.max(lp.utf16End, lp.utf16Start);
      const includeEnd = i === matches.length - 1;
      if (
        charOffset >= lp.utf16Start &&
        (charOffset < end || (includeEnd && charOffset === end))
      ) {
        return lp;
      }
      const distance =
        charOffset < lp.utf16Start
          ? lp.utf16Start - charOffset
          : charOffset > end
            ? charOffset - end
            : 0;
      if (
        distance < nearestDistance ||
        (distance === nearestDistance &&
          (nearest == null ||
            lp.pageIndex > nearest.pageIndex ||
            (lp.pageIndex === nearest.pageIndex && lp.yPt > nearest.yPt)))
      ) {
        nearest = lp;
        nearestDistance = distance;
      }
    }

    return nearest ?? matches[matches.length - 1] ?? null;
  }

  private getCaretRectInParagraph(
    lp: LaidOutParagraph,
    offset: number,
  ): { xPt: number; yPt: number; heightPt: number } | null {
    const paragraphMax = this.getParagraphCharCountByBodyIndex(lp.bodyIndex);
    const minOffset = 0;
    const maxOffset = Math.max(0, paragraphMax);
    const clampedOffset = Math.max(minOffset, Math.min(offset, maxOffset));
    const layoutOffset = clampedOffset + lp.markerUtf16Len;
    const minCaretHeightPt = Math.max(1, DEFAULT_FONT_SIZE_PT * 1.1);

    const trailingAtParagraphEnd =
      clampedOffset === maxOffset && clampedOffset > 0;
    if (trailingAtParagraphEnd) {
      const trailingRect = getFirstRangeRect(
        lp.paragraph,
        this.ck,
        Math.max(lp.markerUtf16Len, layoutOffset - 1),
        layoutOffset,
      );
      if (trailingRect) {
        return {
          xPt: trailingRect[2],
          yPt: trailingRect[1],
          heightPt: Math.max(
            minCaretHeightPt,
            trailingRect[3] - trailingRect[1],
          ),
        };
      }
    }

    const collapsedRect = getFirstRangeRect(
      lp.paragraph,
      this.ck,
      layoutOffset,
      layoutOffset,
    );
    if (collapsedRect) {
      return {
        xPt: collapsedRect[0],
        yPt: collapsedRect[1],
        heightPt: Math.max(
          minCaretHeightPt,
          collapsedRect[3] - collapsedRect[1],
        ),
      };
    }

    if (clampedOffset < maxOffset) {
      const nextRect = getFirstRangeRect(
        lp.paragraph,
        this.ck,
        layoutOffset,
        layoutOffset + 1,
      );
      if (nextRect) {
        return {
          xPt: nextRect[0],
          yPt: nextRect[1],
          heightPt: Math.max(minCaretHeightPt, nextRect[3] - nextRect[1]),
        };
      }
    }

    if (clampedOffset > minOffset) {
      const previousRect = getFirstRangeRect(
        lp.paragraph,
        this.ck,
        Math.max(lp.markerUtf16Len, layoutOffset - 1),
        layoutOffset,
      );
      if (previousRect) {
        return {
          xPt: previousRect[2],
          yPt: previousRect[1],
          heightPt: Math.max(
            minCaretHeightPt,
            previousRect[3] - previousRect[1],
          ),
        };
      }
    }

    // Empty paragraphs can have no glyph rectangles. Fall back to the
    // fragment box so the caret remains visible and typing is possible.
    return {
      xPt: 0,
      yPt: lp.sourceTopPt,
      heightPt: minCaretHeightPt,
    };
  }

  private getParagraphCharCountByBodyIndex(bodyIndex: number): number {
    return this.paragraphCharCountByBodyIndex.get(bodyIndex) ?? 0;
  }

  private hitTestInlineImageInParagraph(
    lp: LaidOutParagraph,
    x: number,
    y: number,
  ): InlineImageHit | null {
    if (!this.currentModel) return null;
    const bodyItem = this.currentModel.body[lp.bodyIndex];
    if (!bodyItem || bodyItem.type !== "paragraph") return null;
    const placeholders = lp.paragraph.getRectsForPlaceholders?.();
    if (!placeholders || placeholders.length === 0) return null;
    const canvas = this.pool.getCanvas(lp.pageIndex);
    const page = this.pages[lp.pageIndex];
    if (!canvas || !page) return null;

    const canvasRect = canvas.getBoundingClientRect();
    const ptToScreenX = canvasRect.width / page.section.pageWidthPt;
    const ptToScreenY = canvasRect.height / page.section.pageHeightPt;
    const originX = page.section.margins.left * ptToScreenX;
    const originY = page.section.margins.top * ptToScreenY;

    let charOffset = 0;
    let placeholderIdx = 0;
    for (const run of bodyItem.runs) {
      if (run.hasTab && !run.inlineImage) {
        placeholderIdx += 1;
      }
      if (run.inlineImage) {
        const placeholder = placeholders[placeholderIdx];
        placeholderIdx += 1;
        if (!placeholder?.rect) {
          charOffset += 1;
          continue;
        }
        const rect = {
          x:
            canvasRect.left +
            originX +
            (lp.xPt + placeholder.rect[0]) * ptToScreenX,
          y:
            canvasRect.top +
            originY +
            (lp.yPt + placeholder.rect[1] - lp.sourceTopPt) * ptToScreenY,
          w: Math.max(
            1,
            (placeholder.rect[2] - placeholder.rect[0]) * ptToScreenX,
          ),
          h: Math.max(
            1,
            (placeholder.rect[3] - placeholder.rect[1]) * ptToScreenY,
          ),
        };
        if (
          x >= rect.x &&
          x <= rect.x + rect.w &&
          y >= rect.y &&
          y <= rect.y + rect.h
        ) {
          return {
            bodyIndex: lp.bodyIndex,
            charOffset,
            imageIndex: run.inlineImage.imageIndex,
            rect,
          };
        }
        charOffset += 1;
        continue;
      }
      if (run.hasBreak || run.hasTab) {
        charOffset += (run.text?.length ?? 0) + 1;
      } else if (run.footnoteRef != null || run.endnoteRef != null) {
        charOffset += String(run.footnoteRef ?? run.endnoteRef ?? "").length;
      } else {
        charOffset += run.text?.length ?? 0;
      }
    }

    return null;
  }

  private inlineImageRectInParagraph(
    lp: LaidOutParagraph,
    targetCharOffset: number,
  ): InlineImageHit | null {
    if (!this.currentModel) return null;
    const bodyItem = this.currentModel.body[lp.bodyIndex];
    if (!bodyItem || bodyItem.type !== "paragraph") return null;
    const placeholders = lp.paragraph.getRectsForPlaceholders?.();
    if (!placeholders || placeholders.length === 0) return null;
    const canvas = this.pool.getCanvas(lp.pageIndex);
    const page = this.pages[lp.pageIndex];
    if (!canvas || !page) return null;

    const canvasRect = canvas.getBoundingClientRect();
    const ptToScreenX = canvasRect.width / page.section.pageWidthPt;
    const ptToScreenY = canvasRect.height / page.section.pageHeightPt;
    const originX = page.section.margins.left * ptToScreenX;
    const originY = page.section.margins.top * ptToScreenY;

    let charOffset = 0;
    let placeholderIdx = 0;
    for (const run of bodyItem.runs) {
      if (run.hasTab && !run.inlineImage) {
        placeholderIdx += 1;
      }
      if (run.inlineImage) {
        const placeholder = placeholders[placeholderIdx];
        placeholderIdx += 1;
        const currentOffset = charOffset;
        charOffset += 1;
        if (currentOffset !== targetCharOffset) {
          continue;
        }
        if (!placeholder?.rect) {
          return null;
        }
        return {
          bodyIndex: lp.bodyIndex,
          charOffset: currentOffset,
          imageIndex: run.inlineImage.imageIndex,
          rect: {
            x:
              canvasRect.left +
              originX +
              (lp.xPt + placeholder.rect[0]) * ptToScreenX,
            y:
              canvasRect.top +
              originY +
              (lp.yPt + placeholder.rect[1] - lp.sourceTopPt) * ptToScreenY,
            w: Math.max(
              1,
              (placeholder.rect[2] - placeholder.rect[0]) * ptToScreenX,
            ),
            h: Math.max(
              1,
              (placeholder.rect[3] - placeholder.rect[1]) * ptToScreenY,
            ),
          },
        };
      }
      if (run.hasBreak || run.hasTab) {
        charOffset += (run.text?.length ?? 0) + 1;
      } else if (run.footnoteRef != null || run.endnoteRef != null) {
        charOffset += String(run.footnoteRef ?? run.endnoteRef ?? "").length;
      } else {
        charOffset += run.text?.length ?? 0;
      }
    }

    return null;
  }

  private rebuildParagraphCharCountCache(model: DocViewModel): void {
    this.paragraphCharCountByBodyIndex.clear();
    for (let i = 0; i < model.body.length; i++) {
      const bodyItem = model.body[i];
      if (!bodyItem || bodyItem.type !== "paragraph") continue;
      this.paragraphCharCountByBodyIndex.set(
        i,
        this.computeParagraphCharCount(bodyItem),
      );
    }
  }

  private computeParagraphCharCount(
    bodyItem: Extract<BodyItem, { type: "paragraph" }>,
  ): number {
    let count = 0;
    for (const run of bodyItem.runs) {
      if (run.inlineImage) {
        count += 1;
      } else if (run.hasBreak || run.hasTab) {
        count += (run.text?.length ?? 0) + 1;
      } else if (run.footnoteRef != null || run.endnoteRef != null) {
        count += String(run.footnoteRef ?? run.endnoteRef ?? "").length;
      } else {
        count += run.text?.length ?? 0;
      }
    }
    return count;
  }

  /**
   * Keep the caret visible while a cursor exists.
   */
  private showCursor(): void {
    this.cursorVisible = true;
    this.updateCaretOverlay();
  }

  /**
   * Hide the caret when selection mode is active.
   */
  private hideCursor(): void {
    this.cursorVisible = false;
    this.hideCaretOverlay();
  }

  private requestPageDraw(pageIndex: number): void {
    if (pageIndex < 0 || pageIndex >= this.pages.length) return;
    if (this.pendingDrawPages.has(pageIndex)) return;

    this.pendingDrawPages.add(pageIndex);
    const epoch = this.renderEpoch;
    const rafId = requestAnimationFrame(() => {
      this.pageDrawRafHandles.delete(pageIndex);
      this.pendingDrawPages.delete(pageIndex);
      if (epoch !== this.renderEpoch) return;
      this.drawPage(pageIndex);
      this.updateCaretOverlay();
    });
    this.pageDrawRafHandles.set(pageIndex, rafId);
  }

  private getSelectionPageIndices(sel: DocSelection | null): Set<number> {
    const pages = new Set<number>();
    if (!sel) return pages;

    const startBody = Math.min(sel.anchor.bodyIndex, sel.focus.bodyIndex);
    const endBody = Math.max(sel.anchor.bodyIndex, sel.focus.bodyIndex);
    for (let bodyIndex = startBody; bodyIndex <= endBody; bodyIndex++) {
      const pageSet = this.paragraphPagesByBodyIndex.get(bodyIndex);
      if (!pageSet) continue;
      for (const pageIndex of pageSet) {
        pages.add(pageIndex);
      }
    }

    return pages;
  }

  private cancelPendingDraws(): void {
    if (this.visibleDrawRaf !== 0) {
      cancelAnimationFrame(this.visibleDrawRaf);
      this.visibleDrawRaf = 0;
    }

    for (const rafId of this.pageDrawRafHandles.values()) {
      cancelAnimationFrame(rafId);
    }
    this.pageDrawRafHandles.clear();
    this.pendingDrawPages.clear();
  }

  private clearMeasurementCaches(): void {
    for (const entry of this.paragraphMeasureCache.values()) {
      if (entry.paragraph?.delete) {
        try {
          entry.paragraph.delete();
        } catch {
          // Ignore cleanup failures.
        }
      }
    }
    this.paragraphMeasureCache.clear();
    this.tableMeasureCache.clear();
  }

  private normalizeCursorPosition(pos: DocPosition): DocPosition {
    const bodyIndex = this.resolveCursorBodyIndex(pos.bodyIndex);
    const maxOffset = this.getParagraphCharCountByBodyIndex(bodyIndex);
    return {
      bodyIndex,
      charOffset: Math.max(0, Math.min(pos.charOffset, maxOffset)),
    };
  }

  private resolveCursorBodyIndex(bodyIndex: number): number {
    if (!this.currentModel || this.currentModel.body.length === 0) {
      return Math.max(0, bodyIndex);
    }

    const clamped = Math.max(
      0,
      Math.min(bodyIndex, this.currentModel.body.length - 1),
    );
    const direct = this.currentModel.body[clamped];
    if (direct?.type === "paragraph") return clamped;

    // Prefer geometry-aware nearest paragraph when available.
    const nearest = this.findNearestParagraphByBodyIndex(clamped);
    if (nearest) return nearest.bodyIndex;

    // Fallback: nearest paragraph by model order.
    let best = -1;
    let bestDistance = Number.POSITIVE_INFINITY;
    for (let i = 0; i < this.currentModel.body.length; i++) {
      if (this.currentModel.body[i]?.type !== "paragraph") continue;
      const distance = Math.abs(i - clamped);
      if (distance < bestDistance) {
        bestDistance = distance;
        best = i;
      }
    }
    return best >= 0 ? best : 0;
  }

  private findNearestParagraphByBodyIndex(
    bodyIndex: number,
  ): LaidOutParagraph | null {
    const exact = this.paragraphsByBodyIndex.get(bodyIndex);
    if (exact && exact.length > 0) return exact[0] ?? null;

    let best: LaidOutParagraph | null = null;
    let bestDistance = Number.POSITIVE_INFINITY;

    for (const [candidateBodyIndex, fragments] of this.paragraphsByBodyIndex) {
      const candidate = fragments[0];
      if (!candidate) continue;
      const distance = Math.abs(candidateBodyIndex - bodyIndex);
      if (
        distance < bestDistance ||
        (distance === bestDistance &&
          (best == null ||
            candidate.bodyIndex < best.bodyIndex ||
            (candidate.bodyIndex === best.bodyIndex &&
              candidate.pageIndex < best.pageIndex)))
      ) {
        best = candidate;
        bestDistance = distance;
      }
    }

    return best;
  }

  private updateCaretOverlay(): void {
    if (!this.cursorPos || this.selection || !this.cursorVisible) {
      this.hideCaretOverlay();
      return;
    }

    const rect = this.getCursorRect();
    if (!rect) {
      this.hideCaretOverlay();
      return;
    }

    const containerRect = this.scrollContainer.getBoundingClientRect();
    const left = rect.x - containerRect.left + this.scrollContainer.scrollLeft;
    const top = rect.y - containerRect.top + this.scrollContainer.scrollTop;
    const rawWidth = Number.isFinite(rect.w) ? rect.w : 1;
    const rawHeight = Number.isFinite(rect.h) ? rect.h : CURSOR_MIN_HEIGHT_PX;
    const finalWidth = Math.max(1, rawWidth);
    const finalHeight = Math.max(CURSOR_MIN_HEIGHT_PX, rawHeight);
    const adjustedTop = top + (rawHeight - finalHeight) / 2;
    const roundedLeft = Math.round(left * 2) / 2;
    const roundedTop = Math.round(adjustedTop * 2) / 2;
    const roundedWidth = Math.round(finalWidth * 2) / 2;
    const roundedHeight = Math.round(finalHeight * 2) / 2;
    const previous = this.lastCaretOverlayRect;
    const unchanged =
      previous != null &&
      previous.left === roundedLeft &&
      previous.top === roundedTop &&
      previous.width === roundedWidth &&
      previous.height === roundedHeight &&
      this.caretOverlay.style.display === "block";
    if (unchanged) return;

    this.caretOverlay.style.left = `${roundedLeft}px`;
    this.caretOverlay.style.top = `${roundedTop}px`;
    this.caretOverlay.style.width = `${roundedWidth}px`;
    this.caretOverlay.style.height = `${roundedHeight}px`;
    this.caretOverlay.style.display = "block";
    this.lastCaretOverlayRect = {
      left: roundedLeft,
      top: roundedTop,
      width: roundedWidth,
      height: roundedHeight,
    };
  }

  private hideCaretOverlay(): void {
    if (this.caretOverlay.style.display === "none") {
      this.lastCaretOverlayRect = null;
      return;
    }
    this.caretOverlay.style.display = "none";
    this.lastCaretOverlayRect = null;
  }

  // ==========================================================================
  // Internal: helpers
  // ==========================================================================

  private hasEditorFocus(): boolean {
    const root = this.scrollContainer.getRootNode() as Document | ShadowRoot;
    const active =
      "activeElement" in root ? root.activeElement : document.activeElement;
    return (
      !!active &&
      (active === this.scrollContainer || this.scrollContainer.contains(active))
    );
  }

  private getPreferredFocusTarget(): HTMLElement | null {
    const canvasInput = this.scrollContainer.querySelector<HTMLElement>(
      'textarea[data-docview-canvas-input="1"]',
    );
    if (canvasInput) return canvasInput;
    return this.dummyInput;
  }

  /**
   * Parse a hex color string (no #) into a CanvasKit color.
   */
  private parseHexColor(hex: string): any {
    const ck = this.ck;
    const clean = hex.replace("#", "");
    const r = parseInt(clean.substring(0, 2), 16) || 0;
    const g = parseInt(clean.substring(2, 4), 16) || 0;
    const b = parseInt(clean.substring(4, 6), 16) || 0;
    return ck.Color(r, g, b, 1.0);
  }

  /**
   * Map alignment string to CanvasKit TextAlign enum.
   */
  private mapAlignment(alignment?: string): any {
    const ck = this.ck;
    switch (alignment) {
      case "center":
        return ck.TextAlign.Center;
      case "right":
        return ck.TextAlign.Right;
      case "justify":
        return ck.TextAlign.Justify;
      default:
        return ck.TextAlign.Left;
    }
  }

  /**
   * Map a Word highlight name to a CanvasKit color.
   */
  private highlightColor(name: string): any {
    const ck = this.ck;
    const map: Record<string, [number, number, number]> = {
      yellow: [255, 255, 0],
      green: [0, 255, 0],
      cyan: [0, 255, 255],
      magenta: [255, 0, 255],
      red: [255, 0, 0],
      blue: [0, 0, 255],
      darkBlue: [0, 0, 128],
      darkCyan: [0, 128, 128],
      darkGreen: [0, 128, 0],
      darkMagenta: [128, 0, 128],
      darkRed: [128, 0, 0],
      darkYellow: [128, 128, 0],
      lightGray: [192, 192, 192],
      darkGray: [128, 128, 128],
      black: [0, 0, 0],
    };
    const rgb = map[name] ?? [255, 255, 0];
    return ck.Color(rgb[0]!, rgb[1]!, rgb[2]!, 1.0);
  }

  /**
   * Decode a data URI into a CanvasKit Image, with caching.
   */
  private getOrDecodeImage(index: number, dataUri: string): any {
    const cached = this.imageCache.get(index);
    if (cached) return cached;

    try {
      const ck = this.ck;
      // Extract base64 data from data URI
      const commaIdx = dataUri.indexOf(",");
      if (commaIdx < 0) return null;
      const base64 = dataUri.substring(commaIdx + 1);
      const binary = atob(base64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      const skImage = ck.MakeImageFromEncoded(bytes);
      if (skImage) {
        this.imageCache.set(index, skImage);
      }
      return skImage;
    } catch {
      return null;
    }
  }

  /** Delete all cached CanvasKit Image objects. */
  private disposeImageCache(): void {
    for (const img of this.imageCache.values()) {
      try {
        img.delete();
      } catch {
        // ignore
      }
    }
    this.imageCache.clear();
  }
}
