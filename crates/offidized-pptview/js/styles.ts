// CSS styles for the PowerPoint viewer.
// Injected into <head> by mount() and into Shadow DOM by <ppt-view>.

export const VIEWER_CSS = `
/* ---- Root container ---- */
.pptview-root {
  font-family: Calibri, "Segoe UI", Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  display: flex;
  height: 100%;
  background: #2b2b2b;
  color: #fff;
  overflow: hidden;
}

/* ---- Filmstrip sidebar ---- */
.pptview-filmstrip {
  width: 160px;
  min-width: 160px;
  overflow-y: auto;
  padding: 12px 8px;
  background: #1e1e1e;
  border-right: 1px solid #3a3a3a;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.pptview-thumbnail-wrapper {
  position: relative;
  cursor: pointer;
  border: 2px solid transparent;
  border-radius: 3px;
  transition: border-color 0.15s;
}

.pptview-thumbnail-wrapper:hover {
  border-color: #555;
}

.pptview-thumbnail-wrapper.active {
  border-color: #4a9eff;
}

.pptview-thumbnail-number {
  position: absolute;
  top: 2px;
  left: 4px;
  font-size: 10px;
  color: #aaa;
  pointer-events: none;
}

.pptview-thumbnail {
  width: 100%;
  aspect-ratio: var(--slide-aspect, 4/3);
  overflow: hidden;
  position: relative;
  background: #fff;
  border-radius: 2px;
}

.pptview-thumbnail-inner {
  transform-origin: top left;
  position: absolute;
  top: 0;
  left: 0;
}

/* ---- Main slide area ---- */
.pptview-slide-area {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: auto;
  padding: 24px;
}

/* ---- Slide canvas ---- */
.pptview-slide {
  background: #fff;
  position: relative;
  overflow: hidden;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.4);
  flex-shrink: 0;
  color: #000;
}

/* ---- Shape (absolutely positioned within slide) ---- */
.pptview-shape {
  position: absolute;
  box-sizing: border-box;
  overflow: hidden;
}

/* ---- Text body within a shape ---- */
.pptview-text-body {
  width: 100%;
  height: 100%;
  box-sizing: border-box;
  overflow: hidden;
  display: flex;
  flex-direction: column;
}

.pptview-text-body[data-anchor="top"] {
  justify-content: flex-start;
}
.pptview-text-body[data-anchor="middle"] {
  justify-content: center;
}
.pptview-text-body[data-anchor="bottom"] {
  justify-content: flex-end;
}

/* ---- Paragraph ---- */
.pptview-paragraph {
  margin: 0;
  padding: 0;
  word-wrap: break-word;
  overflow-wrap: break-word;
}

/* ---- Run ---- */
.pptview-run {
  white-space: pre-wrap;
}

/* ---- Bullet prefix ---- */
.pptview-bullet {
  display: inline-block;
  min-width: 1em;
  text-align: left;
  margin-right: 0.25em;
}

/* ---- Links ---- */
a {
  color: #0563C1;
  text-decoration: underline;
}
a:hover {
  color: #034180;
}

/* ---- Images ---- */
.pptview-image {
  display: block;
  max-width: 100%;
  max-height: 100%;
  object-fit: contain;
}

/* ---- Tables ---- */
.pptview-table {
  border-collapse: collapse;
  width: 100%;
  height: 100%;
}

.pptview-table td {
  padding: 4pt 5pt;
  vertical-align: top;
  text-align: left;
  word-wrap: break-word;
  overflow-wrap: break-word;
  border: 0.5pt solid #d0d0d0;
  font-size: 10pt;
  color: #000;
}

/* ---- Text selection ---- */
::selection {
  background: #b4d5fe;
}
`;
