// ── Feature #5: Table cell formatting ──

use crate::color::ShapeColor;

/// Border style for a single table cell edge.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellBorder {
    /// Line width in EMUs.
    pub width_emu: Option<i64>,
    /// Color as sRGB hex.
    pub color_srgb: Option<String>,
    /// Full color model (supports sRGB, scheme colors, and transforms).
    /// When set, this takes precedence over `color_srgb` during serialization.
    pub color: Option<ShapeColor>,
}

impl CellBorder {
    /// Create a new empty cell border.
    pub fn new() -> Self {
        Self {
            width_emu: None,
            color_srgb: None,
            color: None,
        }
    }

    /// Returns true if any border property is set.
    pub fn is_set(&self) -> bool {
        self.width_emu.is_some() || self.color_srgb.is_some() || self.color.is_some()
    }
}

impl Default for CellBorder {
    fn default() -> Self {
        Self::new()
    }
}

/// Cell border set (top, bottom, left, right, and diagonals).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CellBorders {
    /// Top border (`a:lnT`).
    pub top: Option<CellBorder>,
    /// Bottom border (`a:lnB`).
    pub bottom: Option<CellBorder>,
    /// Left border (`a:lnL`).
    pub left: Option<CellBorder>,
    /// Right border (`a:lnR`).
    pub right: Option<CellBorder>,
    /// Diagonal up border (`a:lnBlToTr` -- bottom-left to top-right).
    pub diagonal_up: Option<CellBorder>,
    /// Diagonal down border (`a:lnTlToBr` -- top-left to bottom-right).
    pub diagonal_down: Option<CellBorder>,
}

impl CellBorders {
    /// Returns true if any border is set.
    pub fn is_set(&self) -> bool {
        self.top.is_some()
            || self.bottom.is_some()
            || self.left.is_some()
            || self.right.is_some()
            || self.diagonal_up.is_some()
            || self.diagonal_down.is_some()
    }
}

/// Vertical text alignment within a table cell (`anchor` attr on `<a:tcPr>`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellTextAnchor {
    Top,
    Middle,
    Bottom,
}

impl CellTextAnchor {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "t" => Some(Self::Top),
            "ctr" => Some(Self::Middle),
            "b" => Some(Self::Bottom),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Top => "t",
            Self::Middle => "ctr",
            Self::Bottom => "b",
        }
    }
}

/// Text direction for table cells (`vert` attr on `<a:tcPr>`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextDirection {
    Horizontal,
    Rotate270,
    Rotate90,
    Stacked,
}

impl TextDirection {
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "horz" => Some(Self::Horizontal),
            "vert270" => Some(Self::Rotate270),
            "vert" => Some(Self::Rotate90),
            "wordArtVert" | "wordArtVertRtl" | "eaVert" => Some(Self::Stacked),
            _ => None,
        }
    }

    pub fn to_xml(self) -> &'static str {
        match self {
            Self::Horizontal => "horz",
            Self::Rotate270 => "vert270",
            Self::Rotate90 => "vert",
            Self::Stacked => "wordArtVert",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TableCell {
    text: String,
    /// Fill color as sRGB hex (from `<a:tcPr>` -> `<a:solidFill>`).
    fill_color_srgb: Option<String>,
    /// Full fill color model (supports sRGB, scheme colors, and transforms).
    fill_color: Option<ShapeColor>,
    /// Cell borders.
    borders: CellBorders,
    /// Bold text formatting (simplified; cell-level override).
    bold: Option<bool>,
    /// Italic text formatting.
    italic: Option<bool>,
    /// Font size in hundredths of a point.
    font_size: Option<u32>,
    /// Font color as sRGB hex.
    font_color_srgb: Option<String>,
    /// Full font color model (supports sRGB, scheme colors, and transforms).
    font_color: Option<ShapeColor>,
    /// Horizontal merge: number of columns this cell spans (`gridSpan` attribute on `a:tc`).
    grid_span: Option<u32>,
    /// Vertical merge: number of rows this cell spans (`rowSpan` attribute on `a:tc`).
    row_span: Option<u32>,
    /// Vertical merge continuation flag (`vMerge="1"` on `a:tc`).
    /// When true, this cell is part of a vertically merged region but not the top cell.
    v_merge: bool,
    /// Vertical text alignment within the cell (`anchor` attr on `<a:tcPr>`).
    vertical_alignment: Option<CellTextAnchor>,
    /// Left margin in EMUs (`marL` attr on `<a:tcPr>`).
    margin_left: Option<i64>,
    /// Right margin in EMUs (`marR` attr on `<a:tcPr>`).
    margin_right: Option<i64>,
    /// Top margin in EMUs (`marT` attr on `<a:tcPr>`).
    margin_top: Option<i64>,
    /// Bottom margin in EMUs (`marB` attr on `<a:tcPr>`).
    margin_bottom: Option<i64>,
    /// Text direction (`vert` attr on `<a:tcPr>`).
    text_direction: Option<TextDirection>,
}

impl TableCell {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into();
    }

    /// Cell fill color as a full `ShapeColor` (supports both sRGB and scheme colors).
    pub fn fill_color(&self) -> Option<&ShapeColor> {
        self.fill_color.as_ref()
    }

    /// Cell fill color as sRGB hex.
    pub fn fill_color_srgb(&self) -> Option<&str> {
        self.fill_color_srgb.as_deref()
    }

    /// Set the full fill color model.
    pub fn set_fill_color(&mut self, color: ShapeColor) {
        self.fill_color = Some(color);
    }

    pub fn set_fill_color_srgb(&mut self, color: impl Into<String>) {
        self.fill_color_srgb = Some(color.into());
    }

    pub fn clear_fill_color_srgb(&mut self) {
        self.fill_color_srgb = None;
    }

    /// Cell borders.
    pub fn borders(&self) -> &CellBorders {
        &self.borders
    }

    pub fn borders_mut(&mut self) -> &mut CellBorders {
        &mut self.borders
    }

    /// Bold text formatting.
    pub fn bold(&self) -> Option<bool> {
        self.bold
    }

    pub fn set_bold(&mut self, bold: bool) {
        self.bold = Some(bold);
    }

    /// Italic text formatting.
    pub fn italic(&self) -> Option<bool> {
        self.italic
    }

    pub fn set_italic(&mut self, italic: bool) {
        self.italic = Some(italic);
    }

    /// Font size in hundredths of a point.
    pub fn font_size(&self) -> Option<u32> {
        self.font_size
    }

    pub fn set_font_size(&mut self, size: u32) {
        self.font_size = Some(size);
    }

    /// Font color as a full `ShapeColor` (supports both sRGB and scheme colors).
    pub fn font_color(&self) -> Option<&ShapeColor> {
        self.font_color.as_ref()
    }

    /// Font color as sRGB hex.
    pub fn font_color_srgb(&self) -> Option<&str> {
        self.font_color_srgb.as_deref()
    }

    /// Set the full font color model.
    pub fn set_font_color(&mut self, color: ShapeColor) {
        self.font_color = Some(color);
    }

    pub fn set_font_color_srgb(&mut self, color: impl Into<String>) {
        self.font_color_srgb = Some(color.into());
    }

    // ── Table merged cells ──

    /// Horizontal merge: number of columns this cell spans (`gridSpan` attribute).
    pub fn grid_span(&self) -> Option<u32> {
        self.grid_span
    }

    /// Set horizontal merge span.
    pub fn set_grid_span(&mut self, span: u32) {
        self.grid_span = Some(span);
    }

    /// Clear horizontal merge span.
    pub fn clear_grid_span(&mut self) {
        self.grid_span = None;
    }

    /// Vertical merge: number of rows this cell spans (`rowSpan` attribute).
    pub fn row_span(&self) -> Option<u32> {
        self.row_span
    }

    /// Set vertical merge span.
    pub fn set_row_span(&mut self, span: u32) {
        self.row_span = Some(span);
    }

    /// Clear vertical merge span.
    pub fn clear_row_span(&mut self) {
        self.row_span = None;
    }

    /// Whether this cell is a vertical merge continuation (`vMerge="1"`).
    pub fn is_v_merge(&self) -> bool {
        self.v_merge
    }

    /// Set vertical merge continuation flag.
    pub fn set_v_merge(&mut self, v_merge: bool) {
        self.v_merge = v_merge;
    }

    // ── Vertical alignment ──

    /// Vertical text alignment within the cell.
    pub fn vertical_alignment(&self) -> Option<CellTextAnchor> {
        self.vertical_alignment
    }

    /// Set vertical text alignment.
    pub fn set_vertical_alignment(&mut self, alignment: CellTextAnchor) {
        self.vertical_alignment = Some(alignment);
    }

    /// Clear vertical text alignment.
    pub fn clear_vertical_alignment(&mut self) {
        self.vertical_alignment = None;
    }

    // ── Cell margins ──

    /// Left margin in EMUs.
    pub fn margin_left(&self) -> Option<i64> {
        self.margin_left
    }

    /// Set left margin in EMUs.
    pub fn set_margin_left(&mut self, emu: i64) {
        self.margin_left = Some(emu);
    }

    /// Clear left margin.
    pub fn clear_margin_left(&mut self) {
        self.margin_left = None;
    }

    /// Right margin in EMUs.
    pub fn margin_right(&self) -> Option<i64> {
        self.margin_right
    }

    /// Set right margin in EMUs.
    pub fn set_margin_right(&mut self, emu: i64) {
        self.margin_right = Some(emu);
    }

    /// Clear right margin.
    pub fn clear_margin_right(&mut self) {
        self.margin_right = None;
    }

    /// Top margin in EMUs.
    pub fn margin_top(&self) -> Option<i64> {
        self.margin_top
    }

    /// Set top margin in EMUs.
    pub fn set_margin_top(&mut self, emu: i64) {
        self.margin_top = Some(emu);
    }

    /// Clear top margin.
    pub fn clear_margin_top(&mut self) {
        self.margin_top = None;
    }

    /// Bottom margin in EMUs.
    pub fn margin_bottom(&self) -> Option<i64> {
        self.margin_bottom
    }

    /// Set bottom margin in EMUs.
    pub fn set_margin_bottom(&mut self, emu: i64) {
        self.margin_bottom = Some(emu);
    }

    /// Clear bottom margin.
    pub fn clear_margin_bottom(&mut self) {
        self.margin_bottom = None;
    }

    // ── Text direction ──

    /// Text direction within the cell.
    pub fn text_direction(&self) -> Option<TextDirection> {
        self.text_direction
    }

    /// Set text direction.
    pub fn set_text_direction(&mut self, direction: TextDirection) {
        self.text_direction = Some(direction);
    }

    /// Clear text direction.
    pub fn clear_text_direction(&mut self) {
        self.text_direction = None;
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Table {
    rows: usize,
    cols: usize,
    cells: Vec<TableCell>,
    /// Column widths in EMUs (Feature #6).
    column_widths_emu: Vec<i64>,
    /// Row heights in EMUs (Feature #6).
    row_heights_emu: Vec<i64>,
    /// Position and size of the table's graphic frame (x, y, cx, cy) in EMUs.
    geometry: Option<(i64, i64, i64, i64)>,
}

impl Table {
    pub fn new(rows: usize, cols: usize) -> Self {
        let cell_count = rows.checked_mul(cols).unwrap_or(0);
        let mut cells = Vec::new();
        cells.resize_with(cell_count, TableCell::new);

        Self {
            rows,
            cols,
            cells,
            column_widths_emu: vec![0; cols],
            row_heights_emu: vec![0; rows],
            geometry: None,
        }
    }

    /// Get the table position and size as (x, y, width, height) in EMUs.
    pub fn geometry(&self) -> Option<(i64, i64, i64, i64)> {
        self.geometry
    }

    /// Set the table position and size in EMUs.
    pub fn set_geometry(&mut self, x: i64, y: i64, cx: i64, cy: i64) {
        self.geometry = Some((x, y, cx, cy));
    }

    pub fn rows(&self) -> usize {
        self.rows
    }

    pub fn cols(&self) -> usize {
        self.cols
    }

    pub fn is_empty(&self) -> bool {
        self.rows == 0 || self.cols == 0
    }

    pub fn cell(&self, row: usize, col: usize) -> Option<&TableCell> {
        let index = self.cell_index(row, col)?;
        self.cells.get(index)
    }

    pub fn cell_mut(&mut self, row: usize, col: usize) -> Option<&mut TableCell> {
        let index = self.cell_index(row, col)?;
        self.cells.get_mut(index)
    }

    pub fn cell_text(&self, row: usize, col: usize) -> Option<&str> {
        self.cell(row, col).map(TableCell::text)
    }

    pub fn set_cell_text(&mut self, row: usize, col: usize, text: impl Into<String>) -> bool {
        let Some(cell) = self.cell_mut(row, col) else {
            return false;
        };
        cell.set_text(text);
        true
    }

    // ── Feature #6: Column widths and row heights ──

    /// Column widths in EMUs.
    pub fn column_widths_emu(&self) -> &[i64] {
        &self.column_widths_emu
    }

    /// Set a column width in EMUs.
    pub fn set_column_width_emu(&mut self, col: usize, width: i64) -> bool {
        if col >= self.cols {
            return false;
        }
        self.column_widths_emu[col] = width;
        true
    }

    /// Row heights in EMUs.
    pub fn row_heights_emu(&self) -> &[i64] {
        &self.row_heights_emu
    }

    /// Set a row height in EMUs.
    pub fn set_row_height_emu(&mut self, row: usize, height: i64) -> bool {
        if row >= self.rows {
            return false;
        }
        self.row_heights_emu[row] = height;
        true
    }

    pub(crate) fn set_column_widths(&mut self, widths: Vec<i64>) {
        self.column_widths_emu = widths;
    }

    pub(crate) fn set_row_heights(&mut self, heights: Vec<i64>) {
        self.row_heights_emu = heights;
    }

    /// Inserts a row at the given index, shifting subsequent rows down.
    /// The new row is populated with default (empty) cells.
    pub fn insert_row(&mut self, at: usize, height_emu: i64) -> bool {
        if at > self.rows {
            return false;
        }
        let insert_pos = at * self.cols;
        for _ in 0..self.cols {
            self.cells.insert(insert_pos, TableCell::new());
        }
        self.row_heights_emu.insert(at, height_emu);
        self.rows += 1;
        true
    }

    /// Removes the row at the given index.
    pub fn remove_row(&mut self, at: usize) -> bool {
        if at >= self.rows || self.rows <= 1 {
            return false;
        }
        let start = at * self.cols;
        let end = start + self.cols;
        if end > self.cells.len() {
            return false;
        }
        self.cells.drain(start..end);
        self.row_heights_emu.remove(at);
        self.rows -= 1;
        true
    }

    /// Inserts a column at the given index, shifting subsequent columns right.
    pub fn insert_column(&mut self, at: usize, width_emu: i64) -> bool {
        if at > self.cols {
            return false;
        }
        // Insert one cell per row at the right position.
        for row in (0..self.rows).rev() {
            let pos = row * self.cols + at;
            self.cells.insert(pos, TableCell::new());
        }
        self.column_widths_emu.insert(at, width_emu);
        self.cols += 1;
        true
    }

    /// Removes the column at the given index.
    pub fn remove_column(&mut self, at: usize) -> bool {
        if at >= self.cols || self.cols <= 1 {
            return false;
        }
        // Remove one cell per row, going in reverse to keep indices stable.
        for row in (0..self.rows).rev() {
            let pos = row * self.cols + at;
            if pos < self.cells.len() {
                self.cells.remove(pos);
            }
        }
        self.column_widths_emu.remove(at);
        self.cols -= 1;
        true
    }

    /// Merge cells in a rectangular region.
    ///
    /// Sets `gridSpan` on the first cell of each row in the region to span columns,
    /// and `rowSpan`/`vMerge` for vertical spanning. Existing content in non-primary
    /// cells is cleared.
    pub fn merge_cells(
        &mut self,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> bool {
        if start_row > end_row
            || start_col > end_col
            || end_row >= self.rows
            || end_col >= self.cols
        {
            return false;
        }

        let col_span = (end_col - start_col + 1) as u32;
        let row_span = (end_row - start_row + 1) as u32;

        for r in start_row..=end_row {
            for c in start_col..=end_col {
                if let Some(cell) = self.cell_mut(r, c) {
                    if r == start_row && c == start_col {
                        // Primary cell: set spans
                        if col_span > 1 {
                            cell.set_grid_span(col_span);
                        }
                        if row_span > 1 {
                            cell.set_row_span(row_span);
                        }
                    } else {
                        // Non-primary cells: clear text and mark as merged
                        cell.set_text("");
                        if c == start_col && r > start_row {
                            // First column of subsequent rows: vertical merge continuation
                            cell.set_v_merge(true);
                            if col_span > 1 {
                                cell.set_grid_span(col_span);
                            }
                        }
                    }
                }
            }
        }
        true
    }

    /// Unmerge all cells in a rectangular region.
    ///
    /// Clears `gridSpan`, `rowSpan`, and `vMerge` attributes from all cells in the region.
    pub fn unmerge_cells(
        &mut self,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> bool {
        if start_row > end_row
            || start_col > end_col
            || end_row >= self.rows
            || end_col >= self.cols
        {
            return false;
        }

        for r in start_row..=end_row {
            for c in start_col..=end_col {
                if let Some(cell) = self.cell_mut(r, c) {
                    cell.clear_grid_span();
                    cell.clear_row_span();
                    cell.set_v_merge(false);
                }
            }
        }
        true
    }

    fn cell_index(&self, row: usize, col: usize) -> Option<usize> {
        if row >= self.rows || col >= self.cols {
            return None;
        }

        row.checked_mul(self.cols)?.checked_add(col)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn set_and_get_cell_text() {
        let mut table = Table::new(2, 2);
        assert!(table.set_cell_text(0, 1, "Quarter"));
        assert_eq!(table.cell_text(0, 1), Some("Quarter"));
    }

    #[test]
    fn bounds_checked_cell_access() {
        let table = Table::new(1, 1);
        assert!(table.cell(0, 0).is_some());
        assert!(table.cell(1, 0).is_none());
        assert!(table.cell(0, 1).is_none());
    }

    #[test]
    fn cell_formatting_defaults() {
        let table = Table::new(1, 1);
        let cell = table.cell(0, 0).unwrap();
        assert!(cell.fill_color_srgb().is_none());
        assert!(cell.bold().is_none());
        assert!(cell.italic().is_none());
        assert!(cell.font_size().is_none());
        assert!(cell.font_color_srgb().is_none());
        assert!(!cell.borders().is_set());
    }

    #[test]
    fn cell_formatting_setters() {
        let mut table = Table::new(1, 1);
        let cell = table.cell_mut(0, 0).unwrap();
        cell.set_fill_color_srgb("AABBCC");
        cell.set_bold(true);
        cell.set_italic(true);
        cell.set_font_size(2400);
        cell.set_font_color_srgb("FF0000");
        cell.borders_mut().top = Some(CellBorder {
            width_emu: Some(12700),
            color_srgb: Some("000000".to_string()),
            color: None,
        });

        assert_eq!(cell.fill_color_srgb(), Some("AABBCC"));
        assert_eq!(cell.bold(), Some(true));
        assert_eq!(cell.italic(), Some(true));
        assert_eq!(cell.font_size(), Some(2400));
        assert_eq!(cell.font_color_srgb(), Some("FF0000"));
        assert!(cell.borders().is_set());
        assert_eq!(cell.borders().top.as_ref().unwrap().width_emu, Some(12700));
    }

    // ── Table merged cells tests ──

    #[test]
    fn cell_grid_span_horizontal_merge() {
        let mut table = Table::new(2, 3);
        let cell = table.cell_mut(0, 0).unwrap();
        assert!(cell.grid_span().is_none());

        cell.set_grid_span(3);
        assert_eq!(cell.grid_span(), Some(3));

        cell.clear_grid_span();
        assert!(cell.grid_span().is_none());
    }

    #[test]
    fn cell_row_span_vertical_merge() {
        let mut table = Table::new(3, 2);
        let cell = table.cell_mut(0, 0).unwrap();
        assert!(cell.row_span().is_none());
        assert!(!cell.is_v_merge());

        cell.set_row_span(2);
        assert_eq!(cell.row_span(), Some(2));

        // The continuation cell below should have v_merge set.
        let continuation = table.cell_mut(1, 0).unwrap();
        continuation.set_v_merge(true);
        assert!(continuation.is_v_merge());

        // Clearing
        let cell = table.cell_mut(0, 0).unwrap();
        cell.clear_row_span();
        assert!(cell.row_span().is_none());

        let continuation = table.cell_mut(1, 0).unwrap();
        continuation.set_v_merge(false);
        assert!(!continuation.is_v_merge());
    }

    #[test]
    fn cell_merge_defaults() {
        let table = Table::new(1, 1);
        let cell = table.cell(0, 0).unwrap();
        assert!(cell.grid_span().is_none());
        assert!(cell.row_span().is_none());
        assert!(!cell.is_v_merge());
    }

    #[test]
    fn column_widths_and_row_heights() {
        let mut table = Table::new(2, 3);
        assert_eq!(table.column_widths_emu().len(), 3);
        assert_eq!(table.row_heights_emu().len(), 2);

        assert!(table.set_column_width_emu(0, 914400));
        assert!(table.set_column_width_emu(1, 1828800));
        assert!(!table.set_column_width_emu(5, 100)); // out of bounds

        assert!(table.set_row_height_emu(0, 457200));
        assert!(!table.set_row_height_emu(5, 100)); // out of bounds

        assert_eq!(table.column_widths_emu()[0], 914400);
        assert_eq!(table.column_widths_emu()[1], 1828800);
        assert_eq!(table.row_heights_emu()[0], 457200);
    }

    // ── Cell text anchor tests ──

    #[test]
    fn cell_text_anchor_xml_roundtrip() {
        for (xml, expected) in [
            ("t", CellTextAnchor::Top),
            ("ctr", CellTextAnchor::Middle),
            ("b", CellTextAnchor::Bottom),
        ] {
            assert_eq!(CellTextAnchor::from_xml(xml), Some(expected));
            assert_eq!(expected.to_xml(), xml);
        }
        assert_eq!(CellTextAnchor::from_xml("unknown"), None);
    }

    // ── Text direction tests ──

    #[test]
    fn text_direction_xml_roundtrip() {
        for (xml, expected, out_xml) in [
            ("horz", TextDirection::Horizontal, "horz"),
            ("vert270", TextDirection::Rotate270, "vert270"),
            ("vert", TextDirection::Rotate90, "vert"),
            ("wordArtVert", TextDirection::Stacked, "wordArtVert"),
        ] {
            assert_eq!(TextDirection::from_xml(xml), Some(expected));
            assert_eq!(expected.to_xml(), out_xml);
        }
        assert_eq!(TextDirection::from_xml("unknown"), None);
    }

    #[test]
    fn text_direction_aliases() {
        assert_eq!(
            TextDirection::from_xml("wordArtVertRtl"),
            Some(TextDirection::Stacked)
        );
        assert_eq!(
            TextDirection::from_xml("eaVert"),
            Some(TextDirection::Stacked)
        );
    }

    // ── Cell vertical alignment tests ──

    #[test]
    fn cell_vertical_alignment_roundtrip() {
        let mut table = Table::new(1, 1);
        let cell = table.cell_mut(0, 0).unwrap();
        assert!(cell.vertical_alignment().is_none());

        cell.set_vertical_alignment(CellTextAnchor::Middle);
        assert_eq!(cell.vertical_alignment(), Some(CellTextAnchor::Middle));

        cell.clear_vertical_alignment();
        assert!(cell.vertical_alignment().is_none());
    }

    // ── Cell margins tests ──

    #[test]
    fn cell_margins_roundtrip() {
        let mut table = Table::new(1, 1);
        let cell = table.cell_mut(0, 0).unwrap();
        assert!(cell.margin_left().is_none());
        assert!(cell.margin_right().is_none());
        assert!(cell.margin_top().is_none());
        assert!(cell.margin_bottom().is_none());

        cell.set_margin_left(91440);
        cell.set_margin_right(91440);
        cell.set_margin_top(45720);
        cell.set_margin_bottom(45720);

        assert_eq!(cell.margin_left(), Some(91440));
        assert_eq!(cell.margin_right(), Some(91440));
        assert_eq!(cell.margin_top(), Some(45720));
        assert_eq!(cell.margin_bottom(), Some(45720));

        cell.clear_margin_left();
        cell.clear_margin_right();
        cell.clear_margin_top();
        cell.clear_margin_bottom();

        assert!(cell.margin_left().is_none());
        assert!(cell.margin_right().is_none());
        assert!(cell.margin_top().is_none());
        assert!(cell.margin_bottom().is_none());
    }

    // ── Cell text direction tests ──

    #[test]
    fn cell_text_direction_roundtrip() {
        let mut table = Table::new(1, 1);
        let cell = table.cell_mut(0, 0).unwrap();
        assert!(cell.text_direction().is_none());

        cell.set_text_direction(TextDirection::Rotate90);
        assert_eq!(cell.text_direction(), Some(TextDirection::Rotate90));

        cell.clear_text_direction();
        assert!(cell.text_direction().is_none());
    }

    // ── Cell defaults include new fields ──

    #[test]
    fn cell_new_fields_default_none() {
        let table = Table::new(1, 1);
        let cell = table.cell(0, 0).unwrap();
        assert!(cell.vertical_alignment().is_none());
        assert!(cell.margin_left().is_none());
        assert!(cell.margin_right().is_none());
        assert!(cell.margin_top().is_none());
        assert!(cell.margin_bottom().is_none());
        assert!(cell.text_direction().is_none());
    }

    // ── Insert/delete row and column tests ──

    #[test]
    fn insert_row_at_end() {
        let mut table = Table::new(2, 2);
        table.set_cell_text(0, 0, "A");
        table.set_cell_text(1, 0, "B");

        assert!(table.insert_row(2, 500000));
        assert_eq!(table.rows(), 3);
        assert_eq!(table.cell_text(0, 0), Some("A"));
        assert_eq!(table.cell_text(1, 0), Some("B"));
        assert_eq!(table.cell_text(2, 0), Some(""));
    }

    #[test]
    fn insert_row_at_beginning() {
        let mut table = Table::new(1, 2);
        table.set_cell_text(0, 0, "X");
        table.set_cell_text(0, 1, "Y");

        assert!(table.insert_row(0, 300000));
        assert_eq!(table.rows(), 2);
        // Old data moved to row 1.
        assert_eq!(table.cell_text(1, 0), Some("X"));
        assert_eq!(table.cell_text(1, 1), Some("Y"));
        // New row 0 is empty.
        assert_eq!(table.cell_text(0, 0), Some(""));
    }

    #[test]
    fn remove_row() {
        let mut table = Table::new(3, 2);
        table.set_cell_text(0, 0, "R0");
        table.set_cell_text(1, 0, "R1");
        table.set_cell_text(2, 0, "R2");

        assert!(table.remove_row(1));
        assert_eq!(table.rows(), 2);
        assert_eq!(table.cell_text(0, 0), Some("R0"));
        assert_eq!(table.cell_text(1, 0), Some("R2"));
    }

    #[test]
    fn remove_last_row_fails() {
        let mut table = Table::new(1, 2);
        assert!(!table.remove_row(0));
    }

    #[test]
    fn insert_column() {
        let mut table = Table::new(2, 2);
        table.set_cell_text(0, 0, "A");
        table.set_cell_text(0, 1, "B");
        table.set_cell_text(1, 0, "C");
        table.set_cell_text(1, 1, "D");

        assert!(table.insert_column(1, 400000));
        assert_eq!(table.cols(), 3);
        assert_eq!(table.cell_text(0, 0), Some("A"));
        assert_eq!(table.cell_text(0, 1), Some("")); // new
        assert_eq!(table.cell_text(0, 2), Some("B"));
        assert_eq!(table.cell_text(1, 0), Some("C"));
        assert_eq!(table.cell_text(1, 1), Some("")); // new
        assert_eq!(table.cell_text(1, 2), Some("D"));
    }

    #[test]
    fn remove_column() {
        let mut table = Table::new(2, 3);
        table.set_cell_text(0, 0, "A");
        table.set_cell_text(0, 1, "B");
        table.set_cell_text(0, 2, "C");

        assert!(table.remove_column(1));
        assert_eq!(table.cols(), 2);
        assert_eq!(table.cell_text(0, 0), Some("A"));
        assert_eq!(table.cell_text(0, 1), Some("C"));
    }

    #[test]
    fn remove_last_column_fails() {
        let mut table = Table::new(2, 1);
        assert!(!table.remove_column(0));
    }

    // ── merge_cells / unmerge_cells tests ──

    #[test]
    fn merge_cells_sets_spans() {
        let mut table = Table::new(3, 4);
        table.set_cell_text(0, 0, "Header");
        table.set_cell_text(0, 1, "Sub1");
        table.set_cell_text(1, 0, "Row1");

        assert!(table.merge_cells(0, 0, 1, 1));

        let primary = table.cell(0, 0).unwrap();
        assert_eq!(primary.grid_span(), Some(2));
        assert_eq!(primary.row_span(), Some(2));

        // Non-primary cells should have text cleared
        assert_eq!(table.cell(0, 1).unwrap().text(), "");
        assert_eq!(table.cell(1, 0).unwrap().text(), "");
    }

    #[test]
    fn merge_cells_single_row_only_gridspan() {
        let mut table = Table::new(2, 4);
        assert!(table.merge_cells(0, 0, 0, 2));

        let primary = table.cell(0, 0).unwrap();
        assert_eq!(primary.grid_span(), Some(3));
        assert!(primary.row_span().is_none()); // no vertical merge for single row
    }

    #[test]
    fn merge_cells_out_of_bounds() {
        let mut table = Table::new(2, 2);
        assert!(!table.merge_cells(0, 0, 5, 5));
        assert!(!table.merge_cells(1, 0, 0, 0)); // start > end
    }

    #[test]
    fn unmerge_cells_clears_merge_attributes() {
        let mut table = Table::new(3, 3);
        table.merge_cells(0, 0, 1, 1);

        assert!(table.unmerge_cells(0, 0, 1, 1));

        let cell = table.cell(0, 0).unwrap();
        assert!(cell.grid_span().is_none());
        assert!(cell.row_span().is_none());
        assert!(!table.cell(1, 0).unwrap().is_v_merge());
    }

    #[test]
    fn unmerge_cells_out_of_bounds() {
        let mut table = Table::new(2, 2);
        assert!(!table.unmerge_cells(0, 0, 5, 5));
    }
}
