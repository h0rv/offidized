// TypeScript interfaces mirroring the Rust view model (model.rs).
// All numeric measurements are in CSS points.

export interface DocViewModel {
  body: BodyItem[];
  sections: SectionModel[];
  images: ImageModel[];
  footnotes: FootnoteModel[];
  endnotes: EndnoteModel[];
}

export type BodyItem =
  | ({ type: "paragraph" } & ParagraphModel)
  | ({ type: "table" } & TableModel);

export interface ParagraphModel {
  runs: RunModel[];
  headingLevel?: number;
  alignment?: string;
  tabStops?: TabStopModel[];
  defaultTabStopPt?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: LineSpacingModel;
  indents?: IndentsModel;
  numbering?: NumberingModel;
  borders?: BordersModel;
  shadingColor?: string;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  sectionIndex: number;
  endsSection?: boolean;
  styleId?: string;
}

export interface RunModel {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  underlineType?: string;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  smallCaps?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  highlight?: string;
  hyperlink?: string;
  hyperlinkTooltip?: string;
  inlineImage?: InlineImageModel;
  floatingImage?: FloatingImageModel;
  footnoteRef?: number;
  endnoteRef?: number;
  hasTab?: boolean;
  hasBreak?: boolean;
}

export interface InlineImageModel {
  imageIndex: number;
  widthPt: number;
  heightPt: number;
  name?: string;
  description?: string;
}

export interface FloatingImageModel {
  imageIndex: number;
  widthPt: number;
  heightPt: number;
  offsetXPt: number;
  offsetYPt: number;
  name?: string;
  description?: string;
  wrapType?: string;
}

export interface TableModel {
  rows: TableRowModel[];
  widthPt?: number;
  alignment?: string;
  columnWidthsPt: number[];
  borders?: BordersModel;
  sectionIndex: number;
}

export interface TableRowModel {
  cells: TableCellModel[];
  heightPt?: number;
  isHeader?: boolean;
  repeatHeader?: boolean;
  keepTogether?: boolean;
  keepWithNext?: boolean;
  cantSplit?: boolean;
}

export interface TableCellModel {
  text: string;
  colSpan?: number;
  rowSpan?: number;
  shadingColor?: string;
  verticalAlign?: string;
  widthPt?: number;
  borders?: BordersModel;
  isCovered?: boolean;
}

export interface SectionModel {
  pageWidthPt: number;
  pageHeightPt: number;
  orientation: string;
  margins: MarginsModel;
  header?: HeaderFooterModel;
  footer?: HeaderFooterModel;
  columnCount?: number;
}

export interface MarginsModel {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface HeaderFooterModel {
  paragraphs: ParagraphModel[];
}

export interface LineSpacingModel {
  value: number;
  rule: string;
}

export interface IndentsModel {
  leftPt?: number;
  rightPt?: number;
  firstLinePt?: number;
  hangingPt?: number;
}

export interface NumberingModel {
  numId: number;
  level: number;
  format: string;
  text: string;
}

export interface BordersModel {
  top?: BorderModel;
  right?: BorderModel;
  bottom?: BorderModel;
  left?: BorderModel;
}

export interface BorderModel {
  style: string;
  color?: string;
  widthPt?: number;
}

export interface ImageModel {
  dataUri: string;
  contentType: string;
}

export interface FootnoteModel {
  id: number;
  text: string;
}

export interface EndnoteModel {
  id: number;
  text: string;
}

export interface TabStopModel {
  positionPt?: number;
  posPt?: number;
  valuePt?: number;
  alignment?: string;
  leader?: string;
}
