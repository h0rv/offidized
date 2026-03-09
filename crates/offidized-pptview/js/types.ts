// TypeScript interfaces mirroring the Rust view model (model.rs).
// All numeric measurements are in CSS points.

export interface PresentationViewModel {
  slides: SlideModel[];
  slideWidthPt: number;
  slideHeightPt: number;
  images: ImageModel[];
}

export interface SlideModel {
  shapes: ShapeModel[];
  background?: BackgroundModel;
  notes?: string;
  hidden?: boolean;
}

export interface ShapeModel {
  xPt: number;
  yPt: number;
  widthPt: number;
  heightPt: number;
  rotation?: number;
  name?: string;
  presetGeometry?: string;
  fill?: ShapeFillModel;
  outline?: OutlineModel;
  text?: TextBodyModel;
  imageIndex?: number;
  table?: TableModel;
  hidden?: boolean;
}

export interface TextBodyModel {
  paragraphs: TextParagraphModel[];
  anchor?: string;
  insets?: InsetsModel;
}

export interface InsetsModel {
  leftPt: number;
  topPt: number;
  rightPt: number;
  bottomPt: number;
}

export interface TextParagraphModel {
  runs: TextRunModel[];
  alignment?: string;
  level?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
  bullet?: BulletModel;
}

export interface TextRunModel {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  hyperlink?: string;
}

export interface BulletModel {
  char?: string;
  autoNumType?: string;
  fontFamily?: string;
  color?: string;
}

export type ShapeFillModel =
  | { type: "solid"; color: string }
  | { type: "gradient"; css: string }
  | { type: "none" };

export interface OutlineModel {
  widthPt?: number;
  color?: string;
  dashStyle?: string;
}

export interface TableModel {
  rows: TableRowModel[];
  columnWidthsPt: number[];
  rowHeightsPt: number[];
}

export interface TableRowModel {
  cells: TableCellModel[];
}

export interface TableCellModel {
  text: string;
  fillColor?: string;
  gridSpan?: number;
  rowSpan?: number;
  vMerge?: boolean;
  verticalAlign?: string;
}

export type BackgroundModel =
  | { type: "solid"; color: string }
  | { type: "gradient"; css: string };

export interface ImageModel {
  dataUri: string;
  contentType: string;
}
