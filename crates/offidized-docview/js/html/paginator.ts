// Simple page break computation for paginated mode.
// Strategy: Render-then-measure approximation.
//
// Moved from ../paginator.ts to html/ since it's DOM-measurement specific.

import type { BodyItem, SectionModel } from "../types.ts";

export interface PageLayout {
  elements: HTMLElement[];
  section: SectionModel | null;
}

/**
 * Split rendered body items into page groups.
 *
 * This is an approximate paginator: it sums element heights and starts a new
 * page when the accumulated height exceeds the available content area, or when
 * a paragraph has `pageBreakBefore`.
 *
 * For true fidelity we'd need to measure mid-paragraph, but for a viewer this
 * is sufficient.
 */
export function paginateItems(
  items: BodyItem[],
  rendered: HTMLElement[],
  sections: SectionModel[],
): PageLayout[] {
  if (rendered.length === 0) {
    return [];
  }

  // We need to measure element heights. Create a hidden measurement container.
  const measurer = document.createElement("div");
  measurer.style.position = "absolute";
  measurer.style.visibility = "hidden";
  measurer.style.left = "-9999px";

  const defaultSection: SectionModel = sections[0] ?? {
    pageWidthPt: 612,
    pageHeightPt: 792,
    orientation: "portrait",
    margins: { top: 72, right: 72, bottom: 72, left: 72 },
  };

  // Set width to content area of first section, with same base styles as viewer
  const contentWidth =
    defaultSection.pageWidthPt -
    defaultSection.margins.left -
    defaultSection.margins.right;
  measurer.style.width = contentWidth + "pt";
  measurer.className = "docview-root";

  document.body.appendChild(measurer);

  const pages: PageLayout[] = [];
  let currentPage: HTMLElement[] = [];
  let currentSection: SectionModel | null =
    sections[items[0]?.sectionIndex ?? 0] ?? defaultSection;

  const availableHeight = () => {
    const sec = currentSection ?? defaultSection;
    return sec.pageHeightPt - sec.margins.top - sec.margins.bottom;
  };

  let usedHeight = 0;

  for (let i = 0; i < rendered.length; i++) {
    const el = rendered[i]!;
    const item = items[i]!;

    // Determine section for this item
    const sectionIdx =
      item.type === "paragraph" ? item.sectionIndex : item.sectionIndex;
    const itemSection = sections[sectionIdx] ?? defaultSection;

    // Page break before?
    const hasPageBreak = item.type === "paragraph" && item.pageBreakBefore;

    if (hasPageBreak && currentPage.length > 0) {
      pages.push({ elements: currentPage, section: currentSection });
      currentPage = [];
      usedHeight = 0;
      currentSection = itemSection;
    }

    // Section change?
    if (currentSection !== itemSection && currentPage.length > 0) {
      pages.push({ elements: currentPage, section: currentSection });
      currentPage = [];
      usedHeight = 0;
      currentSection = itemSection;
    }

    // Measure element height
    measurer.appendChild(el);
    const height = el.getBoundingClientRect().height;
    measurer.removeChild(el);

    // Would this element overflow the page?
    if (usedHeight + height > availableHeight() && currentPage.length > 0) {
      pages.push({ elements: currentPage, section: currentSection });
      currentPage = [];
      usedHeight = 0;
    }

    currentPage.push(el);
    usedHeight += height;
  }

  // Final page
  if (currentPage.length > 0) {
    pages.push({ elements: currentPage, section: currentSection });
  }

  document.body.removeChild(measurer);

  return pages;
}
