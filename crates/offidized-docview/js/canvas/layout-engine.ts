import type { SectionModel } from "../types.ts";

// 1pt = 96/72 CSS px. Internal layout unit is 1/64 CSS px.
const LAYOUT_UNITS_PER_PT = (96 / 72) * 64;

export interface LayoutRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface LayoutMargins {
  t: number;
  r: number;
  b: number;
  l: number;
}

export interface LineLayout {
  lineIndex: number;
  baselineY: number;
  ascent: number;
  descent: number;
  xStart: number;
  xEnd: number;
  range: { startUtf16: number; endUtf16: number };
}

export type BlockFragment =
  | {
      kind: "paragraphFragment";
      bodyIndex: number;
      rect: LayoutRect;
      lineStart: number;
      lineEnd: number;
      utf16Start: number;
      utf16End: number;
      sourceTopPt: number;
      lines: LineLayout[];
      sectionIndex: number;
      columnIndex: number;
      xPt: number;
      yPt: number;
      wPt: number;
      hPt: number;
    }
  | {
      kind: "tableFragment";
      bodyIndex: number;
      rect: LayoutRect;
      rowStart: number;
      rowEnd: number;
      repeatedHeaderRowCount?: number;
      sectionIndex: number;
      columnIndex: number;
      xPt: number;
      yPt: number;
      wPt: number;
      hPt: number;
    };

export interface PageLayout {
  pageIndex: number;
  paper: { w: number; h: number };
  margins: LayoutMargins;
  contentRect: LayoutRect;
  sectionIndex: number;
  columnCount: number;
  fragments: BlockFragment[];
}

export interface LayoutSnapshot {
  docVersion: number;
  pages: PageLayout[];
  hitTest: {
    pages: Array<{
      pageIndex: number;
      fragments: Array<{
        kind: "paragraph" | "table";
        bodyIndex: number;
        rect: LayoutRect;
        utf16Start?: number;
        utf16End?: number;
      }>;
    }>;
  };
  paint: { type: "pageDisplayList"; pageCount: number };
}

export interface ParagraphLineBox {
  lineIndex: number;
  startUtf16: number;
  endUtf16: number;
  topPt: number;
  bottomPt: number;
  baselinePt: number;
  leftPt: number;
  widthPt: number;
}

export interface MeasuredParagraphBlock {
  kind: "paragraph";
  bodyIndex: number;
  sectionIndex: number;
  pageBreakBefore: boolean;
  keepNext: boolean;
  keepLines: boolean;
  spacingBeforePt: number;
  spacingAfterPt: number;
  indentLeftPt: number;
  indentRightPt: number;
  lines: ParagraphLineBox[];
  paragraphWidthPt: number;
}

export interface MeasuredTableBlock {
  kind: "table";
  bodyIndex: number;
  sectionIndex: number;
  rows: Array<{
    heightPt: number;
    isHeader?: boolean;
    repeatHeader?: boolean;
    keepTogether?: boolean;
    keepWithNext?: boolean;
    cantSplit?: boolean;
  }>;
  totalHeightPt: number;
}

export type MeasuredBlock = MeasuredParagraphBlock | MeasuredTableBlock;

export interface LayoutConfig {
  columnGapPt: number;
  widowMinLines: number;
  orphanMinLines: number;
  reservedBottomByPagePt?: number[];
}

export interface LayoutResult {
  snapshot: LayoutSnapshot;
  pages: PageLayout[];
}

interface ColumnGeometry {
  contentWidthPt: number;
  contentHeightPt: number;
  columnCount: number;
  columnGapPt: number;
  columnWidthPt: number;
}

interface CursorState {
  page: PageLayout;
  pageIndex: number;
  sectionIndex: number;
  columnIndex: number;
  usedHeightPt: number;
}

const DEFAULT_LAYOUT_CONFIG: LayoutConfig = {
  columnGapPt: 18,
  widowMinLines: 2,
  orphanMinLines: 2,
};

function ptToLayoutUnits(pt: number): number {
  const raw = pt * LAYOUT_UNITS_PER_PT;
  return raw >= 0 ? Math.floor(raw + 0.5) : Math.ceil(raw - 0.5);
}

function ensurePositive(value: number): number {
  return Number.isFinite(value) && value > 0 ? value : 0;
}

function sectionGeometry(
  section: SectionModel,
  columnGapPt: number,
  reservedBottomPt = 0,
): ColumnGeometry {
  const contentWidthPt =
    section.pageWidthPt - section.margins.left - section.margins.right;
  const contentHeightPt =
    section.pageHeightPt -
    section.margins.top -
    section.margins.bottom -
    Math.max(0, reservedBottomPt);
  const clampedContentWidthPt = ensurePositive(contentWidthPt);
  const columnCount = Math.max(1, Math.floor(section.columnCount ?? 1));
  const totalGapPt = columnGapPt * Math.max(0, columnCount - 1);
  const rawColumnWidth = (clampedContentWidthPt - totalGapPt) / columnCount;
  const columnWidthPt =
    clampedContentWidthPt <= 0 ? 0 : Math.max(1, rawColumnWidth);
  return {
    contentWidthPt: clampedContentWidthPt,
    contentHeightPt: ensurePositive(contentHeightPt),
    columnCount,
    columnGapPt,
    columnWidthPt,
  };
}

function minimumBlockHeightPt(block: MeasuredBlock): number {
  if (block.kind === "table") {
    return block.rows[0]?.heightPt ?? 0;
  }

  const firstLine = block.lines[0];
  if (!firstLine) return block.spacingBeforePt + block.spacingAfterPt;
  const textHeight = Math.max(0, firstLine.bottomPt - firstLine.topPt);
  return block.spacingBeforePt + textHeight;
}

function paragraphTotalHeightPt(block: MeasuredParagraphBlock): number {
  const firstLine = block.lines[0];
  const lastLine = block.lines[block.lines.length - 1];
  const textHeight =
    firstLine && lastLine
      ? Math.max(0, lastLine.bottomPt - firstLine.topPt)
      : 0;
  return block.spacingBeforePt + textHeight + block.spacingAfterPt;
}

function paragraphLineHeightPt(
  block: MeasuredParagraphBlock,
  index: number,
): number {
  const line = block.lines[index];
  if (!line) return 0;
  return Math.max(0, line.bottomPt - line.topPt);
}

function canStartNewRegion(state: CursorState): boolean {
  return state.page.fragments.length > 0 || state.usedHeightPt > 0;
}

function buildPage(
  sectionIndex: number,
  section: SectionModel,
  pageIndex: number,
  reservedBottomPt = 0,
): PageLayout {
  return {
    pageIndex,
    paper: {
      w: ptToLayoutUnits(section.pageWidthPt),
      h: ptToLayoutUnits(section.pageHeightPt),
    },
    margins: {
      t: ptToLayoutUnits(section.margins.top),
      r: ptToLayoutUnits(section.margins.right),
      b: ptToLayoutUnits(section.margins.bottom),
      l: ptToLayoutUnits(section.margins.left),
    },
    contentRect: {
      x: ptToLayoutUnits(section.margins.left),
      y: ptToLayoutUnits(section.margins.top),
      w: ptToLayoutUnits(
        ensurePositive(
          section.pageWidthPt - section.margins.left - section.margins.right,
        ),
      ),
      h: ptToLayoutUnits(
        ensurePositive(
          section.pageHeightPt -
            section.margins.top -
            section.margins.bottom -
            Math.max(0, reservedBottomPt),
        ),
      ),
    },
    sectionIndex,
    columnCount: Math.max(1, Math.floor(section.columnCount ?? 1)),
    fragments: [],
  };
}

function makeParagraphLineLayouts(
  block: MeasuredParagraphBlock,
  start: number,
  end: number,
): LineLayout[] {
  const out: LineLayout[] = [];
  for (let i = start; i < end; i++) {
    const line = block.lines[i];
    if (!line) continue;
    out.push({
      lineIndex: line.lineIndex,
      baselineY: ptToLayoutUnits(line.baselinePt),
      ascent: ptToLayoutUnits(line.topPt - line.baselinePt),
      descent: ptToLayoutUnits(line.bottomPt - line.baselinePt),
      xStart: ptToLayoutUnits(line.leftPt),
      xEnd: ptToLayoutUnits(line.leftPt + line.widthPt),
      range: {
        startUtf16: line.startUtf16,
        endUtf16: line.endUtf16,
      },
    });
  }
  return out;
}

function paragraphSplitCandidate(
  block: MeasuredParagraphBlock,
  startLine: number,
  availableHeightPt: number,
  fullPageHeightPt: number,
  config: LayoutConfig,
): { lineCount: number; strict: boolean } | null {
  const totalLines = block.lines.length;
  if (totalLines === 0 || startLine >= totalLines) return null;

  const remainingLines = totalLines - startLine;

  if (block.keepLines && startLine === 0) {
    const fullHeight = paragraphTotalHeightPt(block);
    if (fullHeight <= availableHeightPt) {
      return { lineCount: remainingLines, strict: true };
    }
    if (fullHeight > fullPageHeightPt) {
      // Taller than a page: allow splitting as a controlled exception.
    } else {
      return null;
    }
  }

  let maxFit = 0;
  for (let count = 1; count <= remainingLines; count++) {
    const endLine = startLine + count;
    const top = block.lines[startLine]?.topPt ?? 0;
    const bottom = block.lines[endLine - 1]?.bottomPt ?? top;
    const textHeight = Math.max(0, bottom - top);
    const spacingBefore = startLine === 0 ? block.spacingBeforePt : 0;
    const spacingAfter = endLine === totalLines ? block.spacingAfterPt : 0;
    const needed = spacingBefore + textHeight + spacingAfter;
    if (needed <= availableHeightPt + 1e-6) {
      maxFit = count;
    } else {
      break;
    }
  }

  if (maxFit <= 0) return null;

  const strictOrphanMin = Math.max(1, config.orphanMinLines);
  const strictWidowMin = Math.max(1, config.widowMinLines);
  const enforceWidowOrphan = totalLines >= strictOrphanMin + strictWidowMin;

  // Strict pass.
  for (let count = maxFit; count >= 1; count--) {
    const endLine = startLine + count;
    const tail = totalLines - endLine;
    if (!enforceWidowOrphan || tail === 0) {
      return { lineCount: count, strict: true };
    }
    if (count >= strictOrphanMin && tail >= strictWidowMin) {
      return { lineCount: count, strict: true };
    }
  }

  // Deterministic relaxed pass.
  return { lineCount: maxFit, strict: false };
}

function pageReservedBottomPt(config: LayoutConfig, pageIndex: number): number {
  const raw = config.reservedBottomByPagePt?.[pageIndex] ?? 0;
  return Number.isFinite(raw) && raw > 0 ? raw : 0;
}

function tableHeaderRowCount(rows: MeasuredTableBlock["rows"]): number {
  let count = 0;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) break;
    if (row.repeatHeader || row.isHeader) {
      count += 1;
      continue;
    }
    break;
  }
  return count;
}

function tableRowsHeightPt(
  rows: MeasuredTableBlock["rows"],
  start: number,
  end: number,
): number {
  let total = 0;
  for (let i = start; i < end; i++) {
    total += rows[i]?.heightPt ?? 0;
  }
  return total;
}

function tableRowGroupEnd(
  rows: MeasuredTableBlock["rows"],
  start: number,
): number {
  let end = Math.min(rows.length, start + 1);
  while (end < rows.length && (rows[end - 1]?.keepWithNext ?? false)) {
    end += 1;
  }
  return end;
}

export function layoutDocument(
  blocks: MeasuredBlock[],
  sections: SectionModel[],
  docVersion: number,
  cfg?: Partial<LayoutConfig>,
): LayoutResult {
  const config: LayoutConfig = { ...DEFAULT_LAYOUT_CONFIG, ...cfg };
  if (blocks.length === 0) {
    const snapshot: LayoutSnapshot = {
      docVersion,
      pages: [],
      hitTest: { pages: [] },
      paint: { type: "pageDisplayList", pageCount: 0 },
    };
    return { snapshot, pages: [] };
  }

  const firstSectionIndex = blocks[0]?.sectionIndex ?? 0;
  const firstSection = sections[firstSectionIndex] ?? sections[0];
  if (!firstSection) {
    const snapshot: LayoutSnapshot = {
      docVersion,
      pages: [],
      hitTest: { pages: [] },
      paint: { type: "pageDisplayList", pageCount: 0 },
    };
    return { snapshot, pages: [] };
  }

  const pages: PageLayout[] = [];
  const pushNewPage = (sectionIndex: number): CursorState => {
    const section = sections[sectionIndex] ?? firstSection;
    const pageIndex = pages.length;
    const reservedBottomPt = pageReservedBottomPt(config, pageIndex);
    const page = buildPage(sectionIndex, section, pageIndex, reservedBottomPt);
    pages.push(page);
    return {
      page,
      pageIndex: page.pageIndex,
      sectionIndex,
      columnIndex: 0,
      usedHeightPt: 0,
    };
  };

  let state = pushNewPage(firstSectionIndex);

  const advanceRegion = (): void => {
    const section = sections[state.sectionIndex] ?? firstSection;
    const geometry = sectionGeometry(
      section,
      config.columnGapPt,
      pageReservedBottomPt(config, state.pageIndex),
    );
    if (state.columnIndex + 1 < geometry.columnCount) {
      state.columnIndex += 1;
      state.usedHeightPt = 0;
      return;
    }
    state = pushNewPage(state.sectionIndex);
  };

  const jumpToSection = (sectionIndex: number): void => {
    if (sectionIndex !== state.sectionIndex) {
      state = pushNewPage(sectionIndex);
      return;
    }
    if (canStartNewRegion(state)) {
      state = pushNewPage(sectionIndex);
    }
  };

  let blockIndex = 0;
  while (blockIndex < blocks.length) {
    const block = blocks[blockIndex]!;
    const sectionIndex = block.sectionIndex;
    const section = sections[sectionIndex] ?? firstSection;

    if (sectionIndex !== state.sectionIndex) {
      state = pushNewPage(sectionIndex);
    }

    const geometryForState = (): ColumnGeometry =>
      sectionGeometry(
        section,
        config.columnGapPt,
        pageReservedBottomPt(config, state.pageIndex),
      );

    if (block.kind === "paragraph" && block.pageBreakBefore) {
      jumpToSection(sectionIndex);
    }

    if (
      block.kind === "paragraph" &&
      block.keepNext &&
      canStartNewRegion(state)
    ) {
      const geometry = geometryForState();
      const available = geometry.contentHeightPt - state.usedHeightPt;
      const thisHeight = paragraphTotalHeightPt(block);
      const nextBlock = blocks[blockIndex + 1];
      const nextMin = nextBlock ? minimumBlockHeightPt(nextBlock) : 0;
      if (thisHeight + nextMin > available + 1e-6) {
        advanceRegion();
        continue;
      }
    }

    if (block.kind === "table") {
      const headerRowCount = tableHeaderRowCount(block.rows);
      let rowStart = 0;
      while (rowStart < block.rows.length) {
        const geometry = geometryForState();
        const available = geometry.contentHeightPt - state.usedHeightPt;
        const repeatedHeaderRowCount = rowStart > 0 ? headerRowCount : 0;
        const repeatedHeaderHeight =
          repeatedHeaderRowCount > 0
            ? tableRowsHeightPt(block.rows, 0, repeatedHeaderRowCount)
            : 0;

        if (
          repeatedHeaderHeight > available + 1e-6 &&
          canStartNewRegion(state)
        ) {
          advanceRegion();
          continue;
        }

        let consumed = repeatedHeaderHeight;
        let rowEnd = rowStart;
        while (rowEnd < block.rows.length) {
          const groupedEnd = tableRowGroupEnd(block.rows, rowEnd);
          const groupHeight = tableRowsHeightPt(block.rows, rowEnd, groupedEnd);
          if (consumed + groupHeight <= available + 1e-6) {
            consumed += groupHeight;
            rowEnd = groupedEnd;
          } else {
            break;
          }
        }

        if (rowEnd === rowStart) {
          if (canStartNewRegion(state)) {
            advanceRegion();
            continue;
          }
          // Overflow row on an empty region.
          const rowH = block.rows[rowStart]?.heightPt ?? 0;
          rowEnd = rowStart + 1;
          consumed = repeatedHeaderHeight + rowH;
        }

        const xPt =
          state.columnIndex * (geometry.columnWidthPt + geometry.columnGapPt);
        const yPt = state.usedHeightPt;
        const rect: LayoutRect = {
          x: ptToLayoutUnits(section.margins.left + xPt),
          y: ptToLayoutUnits(section.margins.top + yPt),
          w: ptToLayoutUnits(geometry.columnWidthPt),
          h: ptToLayoutUnits(consumed),
        };

        state.page.fragments.push({
          kind: "tableFragment",
          bodyIndex: block.bodyIndex,
          rect,
          rowStart,
          rowEnd,
          repeatedHeaderRowCount:
            repeatedHeaderRowCount > 0 ? repeatedHeaderRowCount : undefined,
          sectionIndex,
          columnIndex: state.columnIndex,
          xPt,
          yPt,
          wPt: geometry.columnWidthPt,
          hPt: consumed,
        });

        state.usedHeightPt += consumed;
        rowStart = rowEnd;
        if (rowStart < block.rows.length) {
          advanceRegion();
        }
      }

      blockIndex += 1;
      continue;
    }

    const para = block;
    let lineStart = 0;
    while (lineStart < para.lines.length) {
      const geometry = geometryForState();
      const available = geometry.contentHeightPt - state.usedHeightPt;
      const split = paragraphSplitCandidate(
        para,
        lineStart,
        available,
        geometry.contentHeightPt,
        config,
      );

      if (!split) {
        if (canStartNewRegion(state)) {
          advanceRegion();
          continue;
        }

        // Empty region and no split candidate: force one-line overflow placement.
        const forcedHeight = paragraphLineHeightPt(para, lineStart);
        const forcedEnd = Math.min(para.lines.length, lineStart + 1);
        const top = para.lines[lineStart]?.topPt ?? 0;
        const bottom = para.lines[forcedEnd - 1]?.bottomPt ?? top;
        const textHeight = Math.max(0, bottom - top);
        const spacingBefore = lineStart === 0 ? para.spacingBeforePt : 0;
        const spacingAfter =
          forcedEnd === para.lines.length ? para.spacingAfterPt : 0;
        const placedHeight = spacingBefore + textHeight + spacingAfter;
        const xPt =
          state.columnIndex * (geometry.columnWidthPt + geometry.columnGapPt) +
          para.indentLeftPt;
        const yPt = state.usedHeightPt + spacingBefore;
        const widthPt = Math.max(
          1,
          geometry.columnWidthPt - para.indentLeftPt - para.indentRightPt,
        );
        const rect: LayoutRect = {
          x: ptToLayoutUnits(section.margins.left + xPt),
          y: ptToLayoutUnits(section.margins.top + yPt),
          w: ptToLayoutUnits(widthPt),
          h: ptToLayoutUnits(Math.max(textHeight, forcedHeight)),
        };
        state.page.fragments.push({
          kind: "paragraphFragment",
          bodyIndex: para.bodyIndex,
          rect,
          lineStart,
          lineEnd: forcedEnd,
          utf16Start: para.lines[lineStart]?.startUtf16 ?? 0,
          utf16End: para.lines[forcedEnd - 1]?.endUtf16 ?? 0,
          sourceTopPt: top,
          lines: makeParagraphLineLayouts(para, lineStart, forcedEnd),
          sectionIndex,
          columnIndex: state.columnIndex,
          xPt,
          yPt,
          wPt: widthPt,
          hPt: Math.max(textHeight, forcedHeight),
        });
        state.usedHeightPt += placedHeight;
        lineStart = forcedEnd;
        continue;
      }

      const lineEnd = lineStart + split.lineCount;
      const first = para.lines[lineStart];
      const last = para.lines[lineEnd - 1];
      const top = first?.topPt ?? 0;
      const bottom = last?.bottomPt ?? top;
      const textHeight = Math.max(0, bottom - top);
      const spacingBefore = lineStart === 0 ? para.spacingBeforePt : 0;
      const spacingAfter =
        lineEnd === para.lines.length ? para.spacingAfterPt : 0;
      const placedHeight = spacingBefore + textHeight + spacingAfter;

      if (
        placedHeight > available + 1e-6 &&
        canStartNewRegion(state) &&
        split.lineCount === para.lines.length
      ) {
        advanceRegion();
        continue;
      }

      const xPt =
        state.columnIndex * (geometry.columnWidthPt + geometry.columnGapPt) +
        para.indentLeftPt;
      const yPt = state.usedHeightPt + spacingBefore;
      const widthPt = Math.max(
        1,
        geometry.columnWidthPt - para.indentLeftPt - para.indentRightPt,
      );
      const rect: LayoutRect = {
        x: ptToLayoutUnits(section.margins.left + xPt),
        y: ptToLayoutUnits(section.margins.top + yPt),
        w: ptToLayoutUnits(widthPt),
        h: ptToLayoutUnits(textHeight),
      };

      state.page.fragments.push({
        kind: "paragraphFragment",
        bodyIndex: para.bodyIndex,
        rect,
        lineStart,
        lineEnd,
        utf16Start: first?.startUtf16 ?? 0,
        utf16End: last?.endUtf16 ?? 0,
        sourceTopPt: top,
        lines: makeParagraphLineLayouts(para, lineStart, lineEnd),
        sectionIndex,
        columnIndex: state.columnIndex,
        xPt,
        yPt,
        wPt: widthPt,
        hPt: textHeight,
      });

      state.usedHeightPt += placedHeight;
      lineStart = lineEnd;

      if (lineStart < para.lines.length) {
        advanceRegion();
      }
    }

    blockIndex += 1;
  }

  const snapshot: LayoutSnapshot = {
    docVersion,
    pages,
    hitTest: {
      pages: pages.map((p) => ({
        pageIndex: p.pageIndex,
        fragments: p.fragments.map((f) =>
          f.kind === "paragraphFragment"
            ? {
                kind: "paragraph",
                bodyIndex: f.bodyIndex,
                rect: f.rect,
                utf16Start: f.utf16Start,
                utf16End: f.utf16End,
              }
            : {
                kind: "table",
                bodyIndex: f.bodyIndex,
                rect: f.rect,
              },
        ),
      })),
    },
    paint: {
      type: "pageDisplayList",
      pageCount: pages.length,
    },
  };

  return { snapshot, pages };
}
