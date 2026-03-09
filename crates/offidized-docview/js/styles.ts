// CSS styles for the document viewer.
// Injected into <head> by mount() and into Shadow DOM by <doc-view>.

export const VIEWER_CSS = `
/* ---- Root container ---- */
.docview-root {
  --docview-surface-bg: #f1f3f4;
  --docview-page-bg: #fff;
  --docview-page-shadow: 0 0 0 .75pt #d1d1d1, 0 2px 8px 2px rgba(60,64,67,.16);
  --docview-page-gap: 28px;
  --docview-page-padding-top: 28px;
  --docview-page-padding-bottom: 52px;
  font-family: Calibri, "Segoe UI", Arial, sans-serif;
  font-size: 11pt;
  line-height: 1.15;
  color: #000;
  background: var(--docview-surface-bg);
  /* Prevent inherited antialiased smoothing from making text appear too thin. */
  -webkit-font-smoothing: auto;
  -moz-osx-font-smoothing: auto;
}

/* ---- Continuous mode: single scrolling "page" ---- */
.docview-continuous {
  max-width: var(--page-width, 612pt);
  margin: 24px auto;
  background: var(--docview-page-bg);
  box-shadow: var(--docview-page-shadow);
  padding: var(--margin-top, 72pt) var(--margin-right, 72pt) var(--margin-bottom, 72pt) var(--margin-left, 72pt);
  min-height: 200px;
}

/* ---- Paginated mode: stacked paper pages ---- */
.docview-paginated {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: var(--docview-page-gap);
  padding: var(--docview-page-padding-top) 24px var(--docview-page-padding-bottom);
}

/* ---- Individual page (looks like real paper) ---- */
.docview-page {
  background: var(--docview-page-bg);
  box-shadow: var(--docview-page-shadow);
  position: relative;
  overflow: hidden;
  box-sizing: border-box;
  flex-shrink: 0;
}

.docview-paginated canvas.docview-page {
  user-select: none;
  -webkit-user-select: none;
}

/* Content area within a page (positioned at margin boundaries) */
.docview-page-content {
  position: absolute;
  overflow: hidden;
}

/* ---- Typography ---- */
p, h1, h2, h3, h4, h5, h6 {
  margin: 0;
  padding: 0;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

/* Word Normal style default: 8pt spacing after body paragraphs */
p {
  margin-bottom: 8pt;
}

/* Word built-in heading styles */
h1 { font-size: 16pt; font-weight: bold; color: #2F5496; margin-top: 12pt; margin-bottom: 0; }
h2 { font-size: 13pt; font-weight: bold; color: #2F5496; margin-top: 2pt; margin-bottom: 0; }
h3 { font-size: 12pt; font-weight: bold; color: #1F3763; margin-top: 2pt; margin-bottom: 0; }
h4 { font-size: 11pt; font-weight: bold; font-style: italic; color: #2F5496; margin-top: 2pt; margin-bottom: 0; }
h5 { font-size: 11pt; font-weight: normal; color: #2F5496; margin-top: 2pt; margin-bottom: 0; }
h6 { font-size: 10.5pt; font-weight: normal; font-style: italic; color: #1F3763; margin-top: 2pt; margin-bottom: 0; }

/* ---- Links ---- */
a {
  color: #0563C1;
  text-decoration: underline;
}
a:hover {
  color: #034180;
}

/* ---- Tables ---- */
table {
  border-collapse: collapse;
  table-layout: auto;
  margin: 0;
  max-width: 100%;
}
td, th {
  padding: 0 5.4pt;
  vertical-align: top;
  text-align: left;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

/* ---- Numbering / list bullet prefix ---- */
.docview-numbering {
  display: inline-block;
  min-width: 18pt;
  text-align: right;
  margin-right: 6pt;
}

/* ---- Tab stop ---- */
.docview-tab {
  display: inline-block;
  min-width: 36pt;
}

/* ---- Images ---- */
.docview-inline-image {
  vertical-align: bottom;
  max-width: 100%;
}

.docview-floating-image {
  position: absolute;
}

/* ---- Highlight colors (Word palette) ---- */
.hl-yellow { background-color: #FFFF00; }
.hl-green { background-color: #00FF00; }
.hl-cyan { background-color: #00FFFF; }
.hl-magenta { background-color: #FF00FF; }
.hl-red { background-color: #FF0000; }
.hl-blue { background-color: #0000FF; color: #fff; }
.hl-darkBlue { background-color: #000080; color: #fff; }
.hl-darkCyan { background-color: #008080; }
.hl-darkGreen { background-color: #008000; }
.hl-darkMagenta { background-color: #800080; }
.hl-darkRed { background-color: #800000; color: #fff; }
.hl-darkYellow { background-color: #808000; }
.hl-lightGray { background-color: #C0C0C0; }
.hl-darkGray { background-color: #808080; color: #fff; }
.hl-black { background-color: #000; color: #fff; }

/* ---- Footnotes / Endnotes ---- */
.docview-footnotes, .docview-endnotes {
  border-top: 1px solid #999;
  margin-top: 24pt;
  padding-top: 8pt;
  font-size: 9pt;
  color: #333;
}

.docview-footnote, .docview-endnote {
  margin-bottom: 4pt;
}

.docview-footnote-ref, .docview-endnote-ref {
  font-size: 0.75em;
  vertical-align: super;
  color: #0563C1;
  cursor: pointer;
}

/* ---- Header / Footer in paginated pages ---- */
.docview-page-header,
.docview-page-footer {
  position: absolute;
  left: 0;
  right: 0;
  font-size: 9pt;
  color: #666;
  overflow: hidden;
}

/* ---- Header / Footer in continuous mode ---- */
.docview-header {
  border-bottom: 1px solid #ccc;
  padding-bottom: 6pt;
  margin-bottom: 12pt;
  color: #666;
  font-size: 9pt;
}

.docview-footer {
  border-top: 1px solid #ccc;
  padding-top: 6pt;
  margin-top: 12pt;
  color: #666;
  font-size: 9pt;
}

/* ---- Page break marker (continuous mode) ---- */
.docview-page-break {
  page-break-before: always;
}

/* ---- Text selection ---- */
.docview-root ::selection {
  background: #c2dcff;
}
`;
