// Renderer adapter abstraction for offidized-docview.
//
// Defines shared interfaces that both DOM-based (HTML) and canvas-based
// (CanvasKit) renderers must implement, enabling the editor controller
// to work with either backend.

import type { DocViewModel } from "./types.ts";

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface TableCellPosition {
  bodyIndex: number;
  row: number;
  col: number;
}

export interface TableCellHit extends TableCellPosition {
  rect: Rect;
}

export interface TableCellState extends TableCellPosition {
  rect: Rect | null;
  text: string;
  rowCount: number;
  colCount: number;
}

export interface InlineImageHit {
  bodyIndex: number;
  charOffset: number;
  imageIndex: number;
  rect: Rect;
}

export interface SelectedInlineImageState {
  bodyIndex: number;
  charOffset: number;
  imageIndex: number;
  widthPt: number;
  heightPt: number;
  rect: Rect | null;
}

export interface InlineImageInsertPayload {
  dataUri: string;
  widthPt: number;
  heightPt: number;
  name?: string;
  description?: string;
}

export interface DocPosition {
  bodyIndex: number;
  charOffset: number;
}

export interface DocSelection {
  anchor: DocPosition;
  focus: DocPosition;
}

export interface HitTestResult {
  bodyIndex: number;
  charOffset: number;
  affinity: "leading" | "trailing";
}

export interface RendererAdapter {
  renderModel(model: DocViewModel): void;
  destroy(): void;
  hitTest(x: number, y: number): HitTestResult | null;
  setCursor(pos: DocPosition): void;
  setSelection(sel: DocSelection): void;
  getSelection(): DocSelection | null;
  getCursorRect(): Rect | null;
  getCursorRectForPosition(pos: DocPosition): Rect | null;
  getSelectionRects(sel: DocSelection): Rect[];
  getInlineImageAtPoint?(x: number, y: number): InlineImageHit | null;
  getInlineImageRect?(pos: DocPosition): Rect | null;
  getTableCellAtPoint?(x: number, y: number): TableCellHit | null;
  getTableCellRect?(cell: TableCellPosition): Rect | null;
  getInputElement(): HTMLElement;
  isFocused(): boolean;
  focus(): void;
  getScrollContainer(): HTMLElement;
}

export interface NormalizedInput {
  type:
    | "insertText"
    | "insertFromComposition"
    | "deleteContentBackward"
    | "deleteContentForward"
    | "deleteWordBackward"
    | "deleteWordForward"
    | "insertParagraph"
    | "insertLineBreak"
    | "insertFromPaste"
    | "deleteByCut"
    | "historyUndo"
    | "historyRedo"
    | "insertTab";
  data?: string;
  html?: string;
  shift?: boolean;
}

export type NavigationKey =
  | "ArrowLeft"
  | "ArrowRight"
  | "ArrowUp"
  | "ArrowDown"
  | "Home"
  | "End"
  | "PageUp"
  | "PageDown";

export interface NavigationPayload {
  key: NavigationKey;
  shift: boolean;
  meta: boolean;
  alt: boolean;
  ctrl: boolean;
}

export interface PointerDownPayload {
  x: number;
  y: number;
  clickCount: number;
  shift: boolean;
}

export interface InputAdapter {
  onInput(handler: (input: NormalizedInput) => void): void;
  onShortcut(handler: (key: string, shift: boolean) => void): void;
  onNavigate(handler: (payload: NavigationPayload) => void): void;
  onRequestCopyText(handler: () => string | null | undefined): void;
  onRequestCutText(handler: () => string | null | undefined): void;
  onRequestCopyHtml(handler: () => string | null | undefined): void;
  onRequestCutHtml(handler: () => string | null | undefined): void;
  onPasteImage(handler: (file: File) => void | Promise<void>): void;
  onPointerDown(handler: (payload: PointerDownPayload) => void): void;
  onPointerMove(handler: (x: number, y: number) => void): void;
  onPointerUp(handler: () => void): void;
  onSelectionChange(handler: () => void): void;
  destroy(): void;
}

export type FormatAction = "bold" | "italic" | "underline" | "strikethrough";
export type ListKind = "bullet" | "decimal";
export type ParagraphAlignment = "left" | "center" | "right" | "justify";
export type ImageBlockAlignment = "left" | "center" | "right";

export interface FormattingState {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  highlight?: string;
  hyperlink?: string;
  headingLevel?: number;
  listKind?: ListKind;
  listLevel?: number;
  alignment?: ParagraphAlignment;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  indentLeftPt?: number;
  indentFirstLinePt?: number;
  lineSpacingMultiple?: number;
}

export interface TextStylePatch {
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: boolean | null;
  strike?: boolean | null;
  fontFamily?: string | null;
  fontSizePt?: number | null;
  color?: string | null;
  highlight?: string | null;
  hyperlink?: string | null;
}

export interface ParagraphStylePatch {
  headingLevel?: number | null;
  alignment?: ParagraphAlignment | null;
  spacingBeforePt?: number | null;
  spacingAfterPt?: number | null;
  indentLeftPt?: number | null;
  indentFirstLinePt?: number | null;
  lineSpacingMultiple?: number | null;
  numberingKind?: ListKind | null;
  numberingNumId?: number | null;
  numberingIlvl?: number | null;
}

export interface SyncConfig {
  roomId: string;
  wsUrl?: string;
}

export interface DocEditorController {
  load(data: Uint8Array): number;
  loadBlank(): void;
  replace(data: Uint8Array): number;
  replaceBlank(): void;
  save(): Uint8Array;
  isDirty(): boolean;
  format(action: FormatAction): void;
  setTextStyle(patch: TextStylePatch): void;
  setParagraphStyle(patch: ParagraphStylePatch): void;
  toggleList(kind: ListKind): void;
  insertInlineImage(payload: InlineImageInsertPayload): boolean;
  getSelectedInlineImage(): SelectedInlineImageState | null;
  resizeSelectedInlineImage(widthPt: number, heightPt: number): boolean;
  setSelectedInlineImageAlignment(alignment: ImageBlockAlignment): boolean;
  insertTable(rows: number, columns: number): boolean;
  insertTableRow(): boolean;
  removeTableRow(): boolean;
  insertTableColumn(): boolean;
  removeTableColumn(): boolean;
  getActiveTableCell(): TableCellState | null;
  onActiveTableCellChange(cb: (state: TableCellState | null) => void): void;
  setActiveTableCellText(text: string): boolean;
  moveActiveTableCell(deltaRow: number, deltaCol: number): boolean;
  clearActiveTableCell(): void;
  getFormattingState(): FormattingState;
  onFormattingChange(cb: (state: FormattingState) => void): void;
  getSelectionState(): DocSelection | null;
  getDocEdit(): unknown;
  setAwarenessPausedForTests(paused: boolean): void;
  setTransportPausedForTests(paused: boolean): void;
  destroy(): void;
}
