// DOM renderer: converts PresentationViewModel into styled HTML elements.
// Each slide is a fixed-size canvas with shapes absolutely positioned inside.

import type {
  PresentationViewModel,
  ShapeModel,
  TextBodyModel,
  TextParagraphModel,
  TextRunModel,
  TableModel,
  OutlineModel,
  ShapeFillModel,
} from "./types.ts";

export class SlideRenderer {
  private container: HTMLElement;
  private model: PresentationViewModel | null = null;
  private currentSlide = 0;

  constructor(container: HTMLElement) {
    this.container = container;
  }

  setModel(model: PresentationViewModel): void {
    this.model = model;
    this.currentSlide = 0;
    this.renderSlide(0);
  }

  getSlideCount(): number {
    return this.model?.slides.length ?? 0;
  }

  getCurrentSlide(): number {
    return this.currentSlide;
  }

  goToSlide(index: number): void {
    if (!this.model) return;
    const clamped = Math.max(0, Math.min(index, this.model.slides.length - 1));
    if (clamped !== this.currentSlide) {
      this.currentSlide = clamped;
      this.renderSlide(clamped);
    }
  }

  nextSlide(): void {
    this.goToSlide(this.currentSlide + 1);
  }

  prevSlide(): void {
    this.goToSlide(this.currentSlide - 1);
  }

  renderSlide(index: number): void {
    if (!this.model) return;
    const slide = this.model.slides[index];
    if (!slide) return;

    this.container.innerHTML = "";

    const slideDiv = document.createElement("div");
    slideDiv.className = "pptview-slide";
    slideDiv.style.width = this.model.slideWidthPt + "pt";
    slideDiv.style.height = this.model.slideHeightPt + "pt";

    // Background
    if (slide.background) {
      if (slide.background.type === "solid") {
        slideDiv.style.backgroundColor = "#" + slide.background.color;
      } else if (slide.background.type === "gradient") {
        slideDiv.style.background = slide.background.css;
      }
    }

    // Render each shape
    for (const shape of slide.shapes) {
      if (shape.hidden) continue;
      slideDiv.appendChild(this.renderShape(shape));
    }

    this.container.appendChild(slideDiv);
  }

  // ---------- Slide thumbnails ----------

  renderThumbnail(index: number, thumbnailWidth: number): HTMLElement {
    if (!this.model) return document.createElement("div");
    const slide = this.model.slides[index];
    if (!slide) return document.createElement("div");

    const slideW = this.model.slideWidthPt;
    const slideH = this.model.slideHeightPt;
    const scale = thumbnailWidth / slideW;

    const outer = document.createElement("div");
    outer.className = "pptview-thumbnail";
    outer.style.setProperty("--slide-aspect", `${slideW}/${slideH}`);

    const inner = document.createElement("div");
    inner.className = "pptview-thumbnail-inner";
    inner.style.width = slideW + "pt";
    inner.style.height = slideH + "pt";
    inner.style.transform = `scale(${scale})`;

    // Background
    if (slide.background) {
      if (slide.background.type === "solid") {
        inner.style.backgroundColor = "#" + slide.background.color;
      } else if (slide.background.type === "gradient") {
        inner.style.background = slide.background.css;
      }
    }

    // Render shapes (same as main view)
    for (const shape of slide.shapes) {
      if (shape.hidden) continue;
      inner.appendChild(this.renderShape(shape));
    }

    outer.appendChild(inner);
    return outer;
  }

  // ---------- Shape ----------

  private renderShape(shape: ShapeModel): HTMLElement {
    const div = document.createElement("div");
    div.className = "pptview-shape";

    // Position and size
    div.style.left = shape.xPt + "pt";
    div.style.top = shape.yPt + "pt";
    div.style.width = shape.widthPt + "pt";
    div.style.height = shape.heightPt + "pt";

    // Rotation
    if (shape.rotation) {
      div.style.transform = `rotate(${shape.rotation}deg)`;
    }

    // Fill
    if (shape.fill) {
      applyFill(div, shape.fill);
    }

    // Outline
    if (shape.outline) {
      applyOutline(div, shape.outline);
    }

    // Preset geometry: ellipse gets border-radius
    if (shape.presetGeometry === "ellipse") {
      div.style.borderRadius = "50%";
    } else if (shape.presetGeometry === "roundRect") {
      div.style.borderRadius = "6pt";
    }

    // Image
    if (shape.imageIndex != null && this.model) {
      const imageData = this.model.images[shape.imageIndex];
      if (imageData) {
        const img = document.createElement("img");
        img.className = "pptview-image";
        img.src = imageData.dataUri;
        img.style.width = "100%";
        img.style.height = "100%";
        img.style.objectFit = "contain";
        div.appendChild(img);
      }
    }

    // Table
    if (shape.table) {
      div.appendChild(this.renderTable(shape.table));
    }

    // Text
    if (shape.text) {
      div.appendChild(this.renderTextBody(shape.text));
    }

    return div;
  }

  // ---------- Text body ----------

  private renderTextBody(body: TextBodyModel): HTMLElement {
    const div = document.createElement("div");
    div.className = "pptview-text-body";

    if (body.anchor) {
      div.dataset["anchor"] = body.anchor;
    }

    // Insets as padding
    if (body.insets) {
      div.style.paddingLeft = body.insets.leftPt + "pt";
      div.style.paddingTop = body.insets.topPt + "pt";
      div.style.paddingRight = body.insets.rightPt + "pt";
      div.style.paddingBottom = body.insets.bottomPt + "pt";
    }

    for (const para of body.paragraphs) {
      div.appendChild(this.renderParagraph(para));
    }

    return div;
  }

  // ---------- Paragraph ----------

  private renderParagraph(p: TextParagraphModel): HTMLElement {
    const el = document.createElement("p");
    el.className = "pptview-paragraph";

    // Alignment
    if (p.alignment) {
      el.style.textAlign = p.alignment;
    }

    // Spacing
    if (p.spacingBeforePt != null) {
      el.style.marginTop = p.spacingBeforePt + "pt";
    }
    if (p.spacingAfterPt != null) {
      el.style.marginBottom = p.spacingAfterPt + "pt";
    }

    // Line spacing
    if (p.lineSpacing != null) {
      el.style.lineHeight = String(p.lineSpacing);
    }

    // Indentation from level
    if (p.level != null && p.level > 0) {
      el.style.marginLeft = p.level * 18 + "pt";
    }

    // Bullet
    if (p.bullet) {
      const bulletSpan = document.createElement("span");
      bulletSpan.className = "pptview-bullet";
      if (p.bullet.color) {
        bulletSpan.style.color = "#" + p.bullet.color;
      }
      if (p.bullet.fontFamily) {
        bulletSpan.style.fontFamily = `"${p.bullet.fontFamily}", sans-serif`;
      }
      if (p.bullet.char) {
        bulletSpan.textContent = p.bullet.char;
      } else if (p.bullet.autoNumType) {
        // Auto-numbered bullets: we don't have the counter here,
        // so show a placeholder.
        bulletSpan.textContent = "\u{2022}";
      }
      el.appendChild(bulletSpan);
    }

    // Runs
    for (const run of p.runs) {
      el.appendChild(this.renderRun(run));
    }

    return el;
  }

  // ---------- Run ----------

  private renderRun(r: TextRunModel): HTMLElement {
    const span = document.createElement("span");
    span.className = "pptview-run";
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
    }

    // Strikethrough
    if (r.strikethrough) {
      const existing = span.style.textDecoration;
      span.style.textDecoration = existing
        ? existing + " line-through"
        : "line-through";
    }

    // Color
    if (r.color) {
      span.style.color = "#" + r.color;
    }

    // Hyperlink
    if (r.hyperlink) {
      const a = document.createElement("a");
      a.href = r.hyperlink;
      a.target = "_blank";
      a.rel = "noopener noreferrer";
      a.appendChild(span);
      return a;
    }

    return span;
  }

  // ---------- Table ----------

  private renderTable(t: TableModel): HTMLElement {
    const table = document.createElement("table");
    table.className = "pptview-table";

    for (const row of t.rows) {
      const tr = document.createElement("tr");

      for (const cell of row.cells) {
        if (cell.vMerge) continue;

        const td = document.createElement("td");

        if ((cell.gridSpan ?? 1) > 1) {
          td.colSpan = cell.gridSpan ?? 1;
        }
        if ((cell.rowSpan ?? 1) > 1) {
          td.rowSpan = cell.rowSpan ?? 1;
        }

        if (cell.fillColor) {
          td.style.backgroundColor = "#" + cell.fillColor;
        }
        if (cell.verticalAlign) {
          td.style.verticalAlign = cell.verticalAlign;
        }

        td.textContent = cell.text;
        tr.appendChild(td);
      }

      table.appendChild(tr);
    }

    return table;
  }
}

// ---------- Helpers ----------

function applyFill(el: HTMLElement, fill: ShapeFillModel): void {
  if (fill.type === "solid") {
    el.style.backgroundColor = "#" + fill.color;
  } else if (fill.type === "gradient") {
    el.style.background = fill.css;
  }
  // "none" = transparent, no style needed
}

function applyOutline(el: HTMLElement, outline: OutlineModel): void {
  const width = outline.widthPt != null ? outline.widthPt + "pt" : "1pt";
  const style = outline.dashStyle ?? "solid";
  const color = outline.color ? "#" + outline.color : "#000";
  el.style.border = `${width} ${style} ${color}`;
}
