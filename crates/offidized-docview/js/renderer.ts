// DOM renderer: converts DocViewModel into styled HTML elements.

import type {
  DocViewModel,
  BodyItem,
  ParagraphModel,
  RunModel,
  TableModel,
  SectionModel,
  BordersModel,
  BorderModel,
  HeaderFooterModel,
} from "./types.ts";
import { paginateItems } from "./paginator.ts";

export type RenderMode = "continuous" | "paginated";

export class DocRenderer {
  private container: HTMLElement;
  private model: DocViewModel | null = null;
  private mode: RenderMode = "paginated";

  constructor(container: HTMLElement) {
    this.container = container;
  }

  setModel(model: DocViewModel): void {
    this.model = model;
    this.render();
  }

  setMode(mode: RenderMode): void {
    if (this.mode !== mode) {
      this.mode = mode;
      this.render();
    }
  }

  getMode(): RenderMode {
    return this.mode;
  }

  render(): void {
    if (!this.model) return;
    this.container.innerHTML = "";

    if (this.mode === "continuous") {
      this.renderContinuous(this.model);
    } else {
      this.renderPaginated(this.model);
    }
  }

  // ---------- Continuous mode ----------

  private renderContinuous(model: DocViewModel): void {
    const section = model.sections[0];
    const wrapper = document.createElement("div");
    wrapper.className = "docview-continuous";

    if (section) {
      wrapper.style.setProperty("--page-width", section.pageWidthPt + "pt");
      wrapper.style.setProperty("--margin-top", section.margins.top + "pt");
      wrapper.style.setProperty("--margin-right", section.margins.right + "pt");
      wrapper.style.setProperty(
        "--margin-bottom",
        section.margins.bottom + "pt",
      );
      wrapper.style.setProperty("--margin-left", section.margins.left + "pt");
      wrapper.style.maxWidth = section.pageWidthPt + "pt";
    }

    // Header
    if (section?.header) {
      wrapper.appendChild(this.renderHeaderFooter(section.header, "header"));
    }

    // Body items
    for (const item of model.body) {
      wrapper.appendChild(this.renderBodyItem(item, model));
    }

    // Footnotes
    if (model.footnotes.length > 0) {
      const fnSection = document.createElement("div");
      fnSection.className = "docview-footnotes";
      for (const fn of model.footnotes) {
        const div = document.createElement("div");
        div.className = "docview-footnote";
        div.textContent = `[${fn.id}] ${fn.text}`;
        fnSection.appendChild(div);
      }
      wrapper.appendChild(fnSection);
    }

    // Endnotes
    if (model.endnotes.length > 0) {
      const enSection = document.createElement("div");
      enSection.className = "docview-endnotes";
      for (const en of model.endnotes) {
        const div = document.createElement("div");
        div.className = "docview-endnote";
        div.textContent = `[${en.id}] ${en.text}`;
        enSection.appendChild(div);
      }
      wrapper.appendChild(enSection);
    }

    // Footer
    if (section?.footer) {
      wrapper.appendChild(this.renderHeaderFooter(section.footer, "footer"));
    }

    this.container.appendChild(wrapper);
  }

  // ---------- Paginated mode ----------

  private renderPaginated(model: DocViewModel): void {
    const wrapper = document.createElement("div");
    wrapper.className = "docview-paginated";

    // Pre-render body items to measure heights
    const rendered: HTMLElement[] = model.body.map((item) =>
      this.renderBodyItem(item, model),
    );

    // Use first section for default dimensions
    const defaultSection: SectionModel = model.sections[0] ?? {
      pageWidthPt: 612,
      pageHeightPt: 792,
      orientation: "portrait",
      margins: { top: 72, right: 72, bottom: 72, left: 72 },
    };

    const pages = paginateItems(model.body, rendered, model.sections);

    for (const page of pages) {
      const section = page.section ?? defaultSection;

      // Page container — exact paper dimensions
      const pageDiv = document.createElement("div");
      pageDiv.className = "docview-page";
      pageDiv.style.width = section.pageWidthPt + "pt";
      pageDiv.style.height = section.pageHeightPt + "pt";

      // Header in top margin area
      if (section.header) {
        const headerDiv = this.renderHeaderFooter(section.header, "header");
        headerDiv.className = "docview-page-header";
        headerDiv.style.top = Math.round(section.margins.top * 0.4) + "pt";
        headerDiv.style.left = section.margins.left + "pt";
        headerDiv.style.right = section.margins.right + "pt";
        pageDiv.appendChild(headerDiv);
      }

      // Content area — positioned at margin boundaries, clips overflow
      const contentDiv = document.createElement("div");
      contentDiv.className = "docview-page-content";
      contentDiv.style.top = section.margins.top + "pt";
      contentDiv.style.left = section.margins.left + "pt";
      contentDiv.style.right = section.margins.right + "pt";
      contentDiv.style.bottom = section.margins.bottom + "pt";

      for (const el of page.elements) {
        contentDiv.appendChild(el);
      }
      pageDiv.appendChild(contentDiv);

      // Footer in bottom margin area
      if (section.footer) {
        const footerDiv = this.renderHeaderFooter(section.footer, "footer");
        footerDiv.className = "docview-page-footer";
        footerDiv.style.bottom =
          Math.round(section.margins.bottom * 0.4) + "pt";
        footerDiv.style.left = section.margins.left + "pt";
        footerDiv.style.right = section.margins.right + "pt";
        pageDiv.appendChild(footerDiv);
      }

      wrapper.appendChild(pageDiv);
    }

    this.container.appendChild(wrapper);
  }

  // ---------- Body item ----------

  renderBodyItem(item: BodyItem, model: DocViewModel): HTMLElement {
    if (item.type === "paragraph") {
      return this.renderParagraph(item, model);
    }
    return this.renderTable(item, model);
  }

  // ---------- Paragraph ----------

  renderParagraph(p: ParagraphModel, model: DocViewModel): HTMLElement {
    const level = p.headingLevel;
    const tag =
      level && level >= 1 && level <= 6
        ? (`h${level}` as keyof HTMLElementTagNameMap)
        : "p";
    const el = document.createElement(tag);

    // Alignment
    if (p.alignment) {
      el.style.textAlign = p.alignment;
    }

    // Spacing before
    if (p.spacingBeforePt != null) {
      el.style.marginTop = p.spacingBeforePt + "pt";
    }

    // Spacing after (overrides CSS default of 8pt for <p>)
    if (p.spacingAfterPt != null) {
      el.style.marginBottom = p.spacingAfterPt + "pt";
    }

    // Line spacing
    if (p.lineSpacing) {
      if (p.lineSpacing.rule === "auto") {
        el.style.lineHeight = String(p.lineSpacing.value);
      } else if (p.lineSpacing.rule === "exact") {
        el.style.lineHeight = p.lineSpacing.value + "pt";
      } else {
        // atLeast
        el.style.lineHeight = p.lineSpacing.value + "pt";
      }
    }

    // Indentation
    if (p.indents) {
      if (p.indents.leftPt != null) {
        el.style.marginLeft = p.indents.leftPt + "pt";
      }
      if (p.indents.rightPt != null) {
        el.style.marginRight = p.indents.rightPt + "pt";
      }
      if (p.indents.firstLinePt != null) {
        el.style.textIndent = p.indents.firstLinePt + "pt";
      }
      if (p.indents.hangingPt != null) {
        el.style.textIndent = "-" + p.indents.hangingPt + "pt";
        el.style.paddingLeft = p.indents.hangingPt + "pt";
        const left = p.indents.leftPt ?? 0;
        el.style.marginLeft = left + "pt";
      }
    }

    // Borders
    if (p.borders) {
      applyBorders(el, p.borders);
    }

    // Shading
    if (p.shadingColor) {
      el.style.backgroundColor = "#" + p.shadingColor;
    }

    // Page break
    if (p.pageBreakBefore) {
      el.classList.add("docview-page-break");
    }

    // Numbering prefix
    if (p.numbering) {
      const numSpan = document.createElement("span");
      numSpan.className = "docview-numbering";
      numSpan.textContent = p.numbering.text;
      el.appendChild(numSpan);
    }

    // Runs
    for (const run of p.runs) {
      el.appendChild(this.renderRun(run, model));
    }

    return el;
  }

  // ---------- Run ----------

  renderRun(r: RunModel, model: DocViewModel): HTMLElement | DocumentFragment {
    // Handle line break
    if (r.hasBreak) {
      const frag = document.createDocumentFragment();
      if (r.text) {
        frag.appendChild(this.renderTextRun(r, model));
      }
      frag.appendChild(document.createElement("br"));
      return frag;
    }

    // Handle tab
    if (r.hasTab) {
      const frag = document.createDocumentFragment();
      if (r.text) {
        frag.appendChild(this.renderTextRun(r, model));
      }
      const tabSpan = document.createElement("span");
      tabSpan.className = "docview-tab";
      frag.appendChild(tabSpan);
      return frag;
    }

    // Inline image
    if (r.inlineImage) {
      const img = document.createElement("img");
      img.className = "docview-inline-image";
      img.dataset.docviewInlineImage = "1";
      img.dataset.docviewImageIndex = String(r.inlineImage.imageIndex);
      img.contentEditable = "false";
      const imageData = model.images[r.inlineImage.imageIndex];
      if (imageData) {
        img.src = imageData.dataUri;
      }
      img.style.width = r.inlineImage.widthPt + "pt";
      img.style.height = r.inlineImage.heightPt + "pt";
      if (r.inlineImage.description) {
        img.alt = r.inlineImage.description;
      }
      return img;
    }

    // Floating image
    if (r.floatingImage) {
      const img = document.createElement("img");
      img.className = "docview-floating-image";
      const imageData = model.images[r.floatingImage.imageIndex];
      if (imageData) {
        img.src = imageData.dataUri;
      }
      img.style.width = r.floatingImage.widthPt + "pt";
      img.style.height = r.floatingImage.heightPt + "pt";
      img.style.left = r.floatingImage.offsetXPt + "pt";
      img.style.top = r.floatingImage.offsetYPt + "pt";
      if (r.floatingImage.description) {
        img.alt = r.floatingImage.description;
      }
      return img;
    }

    // Footnote reference
    if (r.footnoteRef != null) {
      const sup = document.createElement("sup");
      sup.className = "docview-footnote-ref";
      sup.textContent = String(r.footnoteRef);
      return sup;
    }

    // Endnote reference
    if (r.endnoteRef != null) {
      const sup = document.createElement("sup");
      sup.className = "docview-endnote-ref";
      sup.textContent = String(r.endnoteRef);
      return sup;
    }

    return this.renderTextRun(r, model);
  }

  private renderTextRun(r: RunModel, _model: DocViewModel): HTMLElement {
    const span = document.createElement("span");
    span.textContent = r.text;

    // Font
    if (r.fontFamily) {
      span.style.fontFamily = `"${r.fontFamily}", Calibri, sans-serif`;
    }
    if (r.fontSizePt != null) {
      span.style.fontSize = r.fontSizePt + "pt";
    }

    // Bold / Italic
    if (r.bold) span.style.fontWeight = "bold";
    if (r.italic) span.style.fontStyle = "italic";

    // Underline
    if (r.underline) {
      span.style.textDecoration = "underline";
      if (r.underlineType && r.underlineType !== "single") {
        if (r.underlineType === "double") {
          span.style.textDecorationStyle = "double";
        } else if (
          r.underlineType === "dotted" ||
          r.underlineType === "dottedheavy"
        ) {
          span.style.textDecorationStyle = "dotted";
        } else if (
          r.underlineType === "dash" ||
          r.underlineType === "dashedheavy" ||
          r.underlineType === "dashlong" ||
          r.underlineType === "dashlongheavy"
        ) {
          span.style.textDecorationStyle = "dashed";
        } else if (
          r.underlineType === "wavy" ||
          r.underlineType === "wavyheavy" ||
          r.underlineType === "wavydouble"
        ) {
          span.style.textDecorationStyle = "wavy";
        }
      }
    }

    // Strikethrough
    if (r.strikethrough) {
      const existing = span.style.textDecoration;
      span.style.textDecoration = existing
        ? existing + " line-through"
        : "line-through";
    }

    // Super/subscript
    if (r.superscript) {
      span.style.verticalAlign = "super";
      span.style.fontSize = "0.75em";
    }
    if (r.subscript) {
      span.style.verticalAlign = "sub";
      span.style.fontSize = "0.75em";
    }

    // Small caps
    if (r.smallCaps) {
      span.style.fontVariant = "small-caps";
    }

    // Color
    if (r.color) {
      span.style.color = "#" + r.color;
    }

    // Highlight
    if (r.highlight) {
      span.style.backgroundColor = this.highlightColor(r.highlight);
    }

    // Hyperlink
    if (r.hyperlink) {
      const a = document.createElement("a");
      a.href = r.hyperlink;
      a.target = "_blank";
      a.rel = "noopener noreferrer";
      if (r.hyperlinkTooltip) {
        a.title = r.hyperlinkTooltip;
      }
      a.appendChild(span);
      return a;
    }

    return span;
  }

  private highlightColor(name: string): string {
    const map: Record<string, string> = {
      yellow: "#ffff00",
      green: "#00ff00",
      cyan: "#00ffff",
      magenta: "#ff00ff",
      red: "#ff0000",
      blue: "#0000ff",
      darkBlue: "#000080",
      darkCyan: "#008080",
      darkGreen: "#008000",
      darkMagenta: "#800080",
      darkRed: "#800000",
      darkYellow: "#808000",
      lightGray: "#c0c0c0",
      darkGray: "#808080",
      black: "#000000",
    };
    return map[name] ?? "#ffff00";
  }

  // ---------- Table ----------

  renderTable(t: TableModel, _model: DocViewModel): HTMLElement {
    const table = document.createElement("table");
    table.style.width = "100%";
    table.style.borderCollapse = "collapse";
    table.style.tableLayout = "fixed";

    const hasExplicitTableBorders = t.borders != null;
    const hasExplicitCellBorders = t.rows.some((row) =>
      row.cells.some((cell) => cell.borders != null),
    );
    const useDefaultGrid = !hasExplicitTableBorders && !hasExplicitCellBorders;

    if (t.alignment) {
      if (t.alignment === "center") {
        table.style.marginLeft = "auto";
        table.style.marginRight = "auto";
      } else if (t.alignment === "right") {
        table.style.marginLeft = "auto";
      }
    }

    if (t.borders) {
      applyBorders(table, t.borders);
    }

    // Compute total cell width from first row for percentage conversion
    const firstRow = t.rows[0];
    let totalCellWidth = 0;
    if (firstRow) {
      for (const cell of firstRow.cells) {
        if (!cell.isCovered) {
          totalCellWidth += cell.widthPt ?? 0;
        }
      }
    }

    for (const [rowIndex, row] of t.rows.entries()) {
      const tr = document.createElement("tr");
      if (row.heightPt != null) {
        tr.style.height = row.heightPt + "pt";
      }

      let colIndex = 0;
      for (const cell of row.cells) {
        if (cell.isCovered) continue;

        const td = document.createElement("td");
        td.dataset.docviewTableCell = "1";
        td.dataset.docviewTableRow = String(rowIndex);
        td.dataset.docviewTableCol = String(colIndex);
        td.contentEditable = "false";

        if ((cell.colSpan ?? 1) > 1) {
          td.colSpan = cell.colSpan ?? 1;
        }
        if ((cell.rowSpan ?? 1) > 1) {
          td.rowSpan = cell.rowSpan ?? 1;
        }

        // Use percentage widths to prevent table overflow
        if (cell.widthPt != null && totalCellWidth > 0) {
          td.style.width =
            ((cell.widthPt / totalCellWidth) * 100).toFixed(2) + "%";
        }
        if (cell.shadingColor) {
          td.style.backgroundColor = "#" + cell.shadingColor;
        }
        if (cell.verticalAlign) {
          td.style.verticalAlign = cell.verticalAlign;
        }
        if (cell.borders) {
          applyBorders(td, cell.borders);
        } else if (useDefaultGrid) {
          td.style.border = "1px solid #c7cdd1";
        }
        td.style.minHeight = "28px";
        td.style.padding = "4px 6px";

        const p = document.createElement("p");
        p.textContent = cell.text;
        td.appendChild(p);

        tr.appendChild(td);
        colIndex += cell.colSpan ?? 1;
      }

      table.appendChild(tr);
    }

    // Wrap in a container to enforce width constraint
    const wrapper = document.createElement("div");
    wrapper.style.overflow = "hidden";
    wrapper.style.width = "100%";
    wrapper.appendChild(table);
    return wrapper;
  }

  // ---------- Header / Footer ----------

  renderHeaderFooter(
    hf: HeaderFooterModel,
    type_: "header" | "footer",
  ): HTMLElement {
    const div = document.createElement("div");
    div.className = type_ === "header" ? "docview-header" : "docview-footer";

    for (const p of hf.paragraphs) {
      div.appendChild(
        this.renderParagraph(p, this.model ?? ({ images: [] } as never)),
      );
    }

    return div;
  }
}

// ---------- Helpers ----------

function applyBorders(el: HTMLElement, borders: BordersModel): void {
  if (borders.top) {
    el.style.borderTop = borderToCss(borders.top);
  }
  if (borders.right) {
    el.style.borderRight = borderToCss(borders.right);
  }
  if (borders.bottom) {
    el.style.borderBottom = borderToCss(borders.bottom);
  }
  if (borders.left) {
    el.style.borderLeft = borderToCss(borders.left);
  }
}

function borderToCss(b: BorderModel): string {
  const width = b.widthPt != null ? b.widthPt + "pt" : "1pt";
  const style = mapBorderStyle(b.style);
  const color = b.color ? "#" + b.color : "#000";
  return `${width} ${style} ${color}`;
}

function mapBorderStyle(style: string): string {
  switch (style) {
    case "single":
      return "solid";
    case "double":
      return "double";
    case "dotted":
      return "dotted";
    case "dashed":
    case "dashSmallGap":
      return "dashed";
    case "triple":
      return "double";
    case "thick":
      return "solid";
    case "wave":
    case "doubleWave":
      return "solid";
    default:
      return "solid";
  }
}
