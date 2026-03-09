# Pagination

How the editor splits content into visual pages.

## DOM Structure

```
#editor-container          (overflow: auto, background: #f0f0f0)
  .docedit-surface         (contenteditable, transparent background, 612pt wide)
    .docedit-page          (white card, 72pt padding, min-height: 792pt)
      p[data-body-index=0]
      p[data-body-index=1]
    .docedit-page-gap      (8px tall, contentEditable="false")
    .docedit-page          (next page card)
      p[data-body-index=2]
      ...
```

The surface is a single `contenteditable` div — this preserves native browser
selection across pages. Page cards provide the white background and shadow.
Gaps are non-editable spacers that let the container's gray background show through.

## Page Dimensions

Letter paper: 8.5 × 11 in = 612 × 792 pt.

- Page card: `min-height: 792pt`, `padding: 72pt` (1-inch margins)
- Content area per page: 792 − 72 − 72 = **648pt** (864px at 96 dpi)
- Content width: 612 − 72 − 72 = **468pt**

## Render Pipeline

On every edit (and on load), `EditRenderer.renderModel()` runs:

1. **Clear and rebuild** — `innerHTML = ""`, create a single `.docedit-page`,
   append all paragraph elements into it.

2. **Force layout** — `void container.offsetHeight` triggers a synchronous
   browser reflow so `getBoundingClientRect()` returns real measurements.

3. **`paginatePages()`** — measures each paragraph's bottom position relative
   to the page content start. When a paragraph's bottom exceeds 648pt (in px),
   a new page group begins. The paragraph that overflows moves to the next page
   (paragraphs are kept whole — no mid-paragraph splitting).

4. **Restructure DOM** — if more than one page group exists, the single page
   is replaced with multiple `.docedit-page` divs separated by `.docedit-page-gap`
   elements.

Measurement happens inside the first page div so paragraphs are rendered at the
correct content width (468pt). Heights measured at the wrong width would give
incorrect page breaks.

## Scroll Preservation

Every full re-render destroys and rebuilds the DOM (`innerHTML = ""`). This
momentarily removes all content, which would clamp the container's `scrollTop`
to 0 and cause a visible jump to page 1.

The fix in `applyAndRender()`:

1. **Save** `container.scrollTop` before calling `renderModel()`
2. **Restore** it immediately after `renderModel()` returns
3. **Focus** the surface with `{ preventScroll: true }` (default `focus()`
   scrolls the element into view, which also jumps to page 1)
4. **`scrollCursorIntoView()`** — uses `Range.getBoundingClientRect()` to
   check if the cursor is within the container's visible area; scrolls just
   enough to show it with 40px of breathing room

Remote updates (`rerenderFromState()`) save/restore scroll but skip cursor
scrolling since the local user's cursor didn't move.

## Click Handling

- **Page gap clicks** — `mousedown` on `.docedit-page-gap` is intercepted,
  cursor placed at end of last paragraph on the page above the gap.
- **Page whitespace clicks** — clicking empty space below the last paragraph
  on a page redirects cursor to the end of that page's last paragraph (checked
  in a `requestAnimationFrame` after the browser places the selection).

## Known Limitations

- **No mid-paragraph splitting** — a paragraph that's taller than a page stays
  on one page (which grows beyond 792pt). Word splits long paragraphs across
  pages by default; we don't yet.
- **No page-aware cursor navigation** — arrow keys near page boundaries rely
  on browser behavior with `contentEditable="false"` gap elements. This mostly
  works but can occasionally cause the cursor to skip a line.
- **Fixed page size** — always letter (612 × 792pt). No support for A4 or
  custom paper sizes yet.
- **No headers/footers** — page cards are pure content areas with no
  header/footer rendering.
