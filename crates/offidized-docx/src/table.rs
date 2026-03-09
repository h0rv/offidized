use offidized_opc::RawXmlNode;

/// A single border edge definition for a table border model.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TableBorder {
    line_type: Option<String>,
    size_eighth_points: Option<u16>,
    color: Option<String>,
    space_eighth_points: Option<u16>,
}

impl TableBorder {
    pub fn new(line_type: impl Into<String>) -> Self {
        let mut border = Self::default();
        border.set_line_type(line_type);
        border
    }

    pub fn line_type(&self) -> Option<&str> {
        self.line_type.as_deref()
    }

    pub fn set_line_type(&mut self, line_type: impl Into<String>) {
        self.line_type = normalize_optional_text(line_type.into());
    }

    pub fn clear_line_type(&mut self) {
        self.line_type = None;
    }

    pub fn size_eighth_points(&self) -> Option<u16> {
        self.size_eighth_points
    }

    pub fn set_size_eighth_points(&mut self, size: u16) {
        self.size_eighth_points = Some(size);
    }

    pub fn clear_size_eighth_points(&mut self) {
        self.size_eighth_points = None;
    }

    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    pub fn set_color(&mut self, color: impl Into<String>) {
        self.color = normalize_color_value(color.into().as_str());
    }

    pub fn clear_color(&mut self) {
        self.color = None;
    }

    /// Space between border and content, in eighth-points (`w:space`).
    pub fn space_eighth_points(&self) -> Option<u16> {
        self.space_eighth_points
    }

    /// Set space between border and content, in eighth-points.
    pub fn set_space_eighth_points(&mut self, space: u16) {
        self.space_eighth_points = Some(space);
    }

    /// Clear space.
    pub fn clear_space_eighth_points(&mut self) {
        self.space_eighth_points = None;
    }

    pub(crate) fn set_line_type_option(&mut self, line_type: Option<String>) {
        self.line_type = line_type.and_then(normalize_optional_text);
    }
}

/// Table border collection (`w:tblBorders`).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TableBorders {
    top: Option<TableBorder>,
    left: Option<TableBorder>,
    bottom: Option<TableBorder>,
    right: Option<TableBorder>,
    inside_horizontal: Option<TableBorder>,
    inside_vertical: Option<TableBorder>,
}

impl TableBorders {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn top(&self) -> Option<&TableBorder> {
        self.top.as_ref()
    }

    pub fn top_mut(&mut self) -> &mut Option<TableBorder> {
        &mut self.top
    }

    pub fn set_top(&mut self, border: TableBorder) {
        self.top = Some(border);
    }

    pub fn left(&self) -> Option<&TableBorder> {
        self.left.as_ref()
    }

    pub fn left_mut(&mut self) -> &mut Option<TableBorder> {
        &mut self.left
    }

    pub fn set_left(&mut self, border: TableBorder) {
        self.left = Some(border);
    }

    pub fn bottom(&self) -> Option<&TableBorder> {
        self.bottom.as_ref()
    }

    pub fn bottom_mut(&mut self) -> &mut Option<TableBorder> {
        &mut self.bottom
    }

    pub fn set_bottom(&mut self, border: TableBorder) {
        self.bottom = Some(border);
    }

    pub fn right(&self) -> Option<&TableBorder> {
        self.right.as_ref()
    }

    pub fn right_mut(&mut self) -> &mut Option<TableBorder> {
        &mut self.right
    }

    pub fn set_right(&mut self, border: TableBorder) {
        self.right = Some(border);
    }

    pub fn inside_horizontal(&self) -> Option<&TableBorder> {
        self.inside_horizontal.as_ref()
    }

    pub fn inside_horizontal_mut(&mut self) -> &mut Option<TableBorder> {
        &mut self.inside_horizontal
    }

    pub fn set_inside_horizontal(&mut self, border: TableBorder) {
        self.inside_horizontal = Some(border);
    }

    pub fn inside_vertical(&self) -> Option<&TableBorder> {
        self.inside_vertical.as_ref()
    }

    pub fn inside_vertical_mut(&mut self) -> &mut Option<TableBorder> {
        &mut self.inside_vertical
    }

    pub fn set_inside_vertical(&mut self, border: TableBorder) {
        self.inside_vertical = Some(border);
    }

    pub fn clear(&mut self) {
        *self = Self::default();
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.top.is_none()
            && self.left.is_none()
            && self.bottom.is_none()
            && self.right.is_none()
            && self.inside_horizontal.is_none()
            && self.inside_vertical.is_none()
    }
}

/// Vertical merge mode for a table cell (`w:vMerge`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalMerge {
    /// This cell starts a new vertical merge group (`w:vMerge w:val="restart"`).
    Restart,
    /// This cell continues a vertical merge from the cell above (`w:vMerge` with no val).
    Continue,
}

/// Vertical text alignment within a table cell (`w:vAlign`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlignment {
    /// Top alignment (`w:vAlign w:val="top"`).
    Top,
    /// Center alignment (`w:vAlign w:val="center"`).
    Center,
    /// Bottom alignment (`w:vAlign w:val="bottom"`).
    Bottom,
}

/// Cell border collection (`w:tcBorders`).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CellBorders {
    top: Option<TableBorder>,
    left: Option<TableBorder>,
    bottom: Option<TableBorder>,
    right: Option<TableBorder>,
}

impl CellBorders {
    /// Create an empty cell border set.
    pub fn new() -> Self {
        Self::default()
    }

    /// Top border.
    pub fn top(&self) -> Option<&TableBorder> {
        self.top.as_ref()
    }

    /// Set top border.
    pub fn set_top(&mut self, border: TableBorder) {
        self.top = Some(border);
    }

    /// Clear top border.
    pub fn clear_top(&mut self) {
        self.top = None;
    }

    /// Left border.
    pub fn left(&self) -> Option<&TableBorder> {
        self.left.as_ref()
    }

    /// Set left border.
    pub fn set_left(&mut self, border: TableBorder) {
        self.left = Some(border);
    }

    /// Clear left border.
    pub fn clear_left(&mut self) {
        self.left = None;
    }

    /// Bottom border.
    pub fn bottom(&self) -> Option<&TableBorder> {
        self.bottom.as_ref()
    }

    /// Set bottom border.
    pub fn set_bottom(&mut self, border: TableBorder) {
        self.bottom = Some(border);
    }

    /// Clear bottom border.
    pub fn clear_bottom(&mut self) {
        self.bottom = None;
    }

    /// Right border.
    pub fn right(&self) -> Option<&TableBorder> {
        self.right.as_ref()
    }

    /// Set right border.
    pub fn set_right(&mut self, border: TableBorder) {
        self.right = Some(border);
    }

    /// Clear right border.
    pub fn clear_right(&mut self) {
        self.right = None;
    }

    /// Clear all borders.
    pub fn clear(&mut self) {
        *self = Self::default();
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.top.is_none() && self.left.is_none() && self.bottom.is_none() && self.right.is_none()
    }
}

/// Cell margin/padding values in twips.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct CellMargins {
    top_twips: Option<u32>,
    left_twips: Option<u32>,
    bottom_twips: Option<u32>,
    right_twips: Option<u32>,
}

impl CellMargins {
    /// Create empty cell margins.
    pub fn new() -> Self {
        Self::default()
    }

    /// Top margin in twips.
    pub fn top_twips(&self) -> Option<u32> {
        self.top_twips
    }

    /// Set top margin in twips.
    pub fn set_top_twips(&mut self, value: u32) {
        self.top_twips = Some(value);
    }

    /// Left margin in twips.
    pub fn left_twips(&self) -> Option<u32> {
        self.left_twips
    }

    /// Set left margin in twips.
    pub fn set_left_twips(&mut self, value: u32) {
        self.left_twips = Some(value);
    }

    /// Bottom margin in twips.
    pub fn bottom_twips(&self) -> Option<u32> {
        self.bottom_twips
    }

    /// Set bottom margin in twips.
    pub fn set_bottom_twips(&mut self, value: u32) {
        self.bottom_twips = Some(value);
    }

    /// Right margin in twips.
    pub fn right_twips(&self) -> Option<u32> {
        self.right_twips
    }

    /// Set right margin in twips.
    pub fn set_right_twips(&mut self, value: u32) {
        self.right_twips = Some(value);
    }

    /// Clear all margins.
    pub fn clear(&mut self) {
        *self = Self::default();
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.top_twips.is_none()
            && self.left_twips.is_none()
            && self.bottom_twips.is_none()
            && self.right_twips.is_none()
    }
}

/// Row-level properties for a table row.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TableRowProperties {
    height_twips: Option<u32>,
    /// Height rule: `"exact"` or `"atLeast"`.
    height_rule: Option<String>,
    /// Whether this row should repeat as a header row at the top of each page.
    repeat_header: bool,
}

impl TableRowProperties {
    /// Create empty row properties.
    pub fn new() -> Self {
        Self::default()
    }

    /// Row height in twips.
    pub fn height_twips(&self) -> Option<u32> {
        self.height_twips
    }

    /// Set row height in twips.
    pub fn set_height_twips(&mut self, value: u32) {
        self.height_twips = Some(value);
    }

    /// Clear row height.
    pub fn clear_height_twips(&mut self) {
        self.height_twips = None;
    }

    /// Row height rule (`"exact"` or `"atLeast"`).
    pub fn height_rule(&self) -> Option<&str> {
        self.height_rule.as_deref()
    }

    /// Set row height rule.
    pub fn set_height_rule(&mut self, rule: impl Into<String>) {
        self.height_rule = Some(rule.into());
    }

    /// Clear row height rule.
    pub fn clear_height_rule(&mut self) {
        self.height_rule = None;
    }

    /// Whether this row should repeat as a header at the top of each page.
    pub fn repeat_header(&self) -> bool {
        self.repeat_header
    }

    /// Set repeat header row flag.
    pub fn set_repeat_header(&mut self, repeat: bool) {
        self.repeat_header = repeat;
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.height_twips.is_none() && self.height_rule.is_none() && !self.repeat_header
    }
}

/// A single table cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableCell {
    text: String,
    horizontal_span: usize,
    horizontal_merge_continuation: bool,
    vertical_merge: Option<VerticalMerge>,
    shading_color: Option<String>,
    shading_color_attribute: Option<String>,
    shading_pattern: Option<String>,
    vertical_alignment: Option<VerticalAlignment>,
    cell_width_twips: Option<u32>,
    borders: CellBorders,
    margins: CellMargins,
    unknown_property_children: Vec<RawXmlNode>,
}

impl Default for TableCell {
    fn default() -> Self {
        Self {
            text: String::new(),
            horizontal_span: 1,
            horizontal_merge_continuation: false,
            vertical_merge: None,
            shading_color: None,
            shading_color_attribute: None,
            shading_pattern: None,
            vertical_alignment: None,
            cell_width_twips: None,
            borders: CellBorders::new(),
            margins: CellMargins::new(),
            unknown_property_children: Vec::new(),
        }
    }
}

impl TableCell {
    /// Create an empty cell.
    pub fn new() -> Self {
        Self::default()
    }

    /// Get the cell text.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Replace the cell text.
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into();
    }

    /// Horizontal grid span value for this cell (`w:gridSpan`).
    pub fn horizontal_span(&self) -> usize {
        self.horizontal_span
    }

    /// Whether this cell is covered by a previous horizontally merged cell.
    pub fn is_horizontal_merge_continuation(&self) -> bool {
        self.horizontal_merge_continuation
    }

    /// Vertical merge mode (`w:vMerge`).
    pub fn vertical_merge(&self) -> Option<VerticalMerge> {
        self.vertical_merge
    }

    /// Set vertical merge mode.
    pub fn set_vertical_merge(&mut self, vertical_merge: VerticalMerge) {
        self.vertical_merge = Some(vertical_merge);
    }

    /// Clear vertical merge.
    pub fn clear_vertical_merge(&mut self) {
        self.vertical_merge = None;
    }

    /// Shading fill color (`w:shd w:fill`), uppercase hex without `#`.
    pub fn shading_color(&self) -> Option<&str> {
        self.shading_color.as_deref()
    }

    /// Set shading fill color. Normalizes by stripping `#` and uppercasing.
    pub fn set_shading_color(&mut self, color: impl Into<String>) {
        self.shading_color = normalize_color_value(color.into().as_str());
    }

    /// Clear shading fill color.
    pub fn clear_shading_color(&mut self) {
        self.shading_color = None;
    }

    /// The `w:color` attribute on cell shading (`w:shd`).
    ///
    /// This is the foreground/pattern color, not the fill. Returns `None` when the
    /// attribute was not present in the original XML.
    pub fn shading_color_attribute(&self) -> Option<&str> {
        self.shading_color_attribute.as_deref()
    }

    /// Set the `w:color` attribute for cell shading.
    pub fn set_shading_color_attribute(&mut self, color: impl Into<String>) {
        let color = color.into();
        self.shading_color_attribute = if color.trim().is_empty() {
            None
        } else {
            Some(color)
        };
    }

    /// Clear the shading `w:color` attribute (defaults to `"auto"` on serialize).
    pub fn clear_shading_color_attribute(&mut self) {
        self.shading_color_attribute = None;
    }

    /// Shading pattern value (`w:shd w:val`), e.g. `"clear"`, `"diagStripe"`.
    pub fn shading_pattern(&self) -> Option<&str> {
        self.shading_pattern.as_deref()
    }

    /// Set the shading pattern value.
    pub fn set_shading_pattern(&mut self, pattern: impl Into<String>) {
        let pattern = pattern.into();
        self.shading_pattern = if pattern.trim().is_empty() {
            None
        } else {
            Some(pattern)
        };
    }

    /// Set shading pattern from an `Option<String>` (used during parse).
    pub(crate) fn set_shading_pattern_option(&mut self, pattern: Option<String>) {
        self.shading_pattern = pattern;
    }

    /// Clear shading pattern value.
    pub fn clear_shading_pattern(&mut self) {
        self.shading_pattern = None;
    }

    /// Vertical text alignment within the cell (`w:vAlign`).
    pub fn vertical_alignment(&self) -> Option<VerticalAlignment> {
        self.vertical_alignment
    }

    /// Set vertical text alignment.
    pub fn set_vertical_alignment(&mut self, alignment: VerticalAlignment) {
        self.vertical_alignment = Some(alignment);
    }

    /// Clear vertical text alignment.
    pub fn clear_vertical_alignment(&mut self) {
        self.vertical_alignment = None;
    }

    /// Cell width in twips (`w:tcW w:type="dxa"`).
    pub fn cell_width_twips(&self) -> Option<u32> {
        self.cell_width_twips
    }

    /// Set cell width in twips.
    pub fn set_cell_width_twips(&mut self, width: u32) {
        self.cell_width_twips = Some(width);
    }

    /// Clear cell width.
    pub fn clear_cell_width_twips(&mut self) {
        self.cell_width_twips = None;
    }

    /// Cell borders (`w:tcBorders`).
    pub fn borders(&self) -> &CellBorders {
        &self.borders
    }

    /// Mutable cell borders.
    pub fn borders_mut(&mut self) -> &mut CellBorders {
        &mut self.borders
    }

    /// Set cell borders.
    pub fn set_borders(&mut self, borders: CellBorders) {
        self.borders = borders;
    }

    /// Clear all cell borders.
    pub fn clear_borders(&mut self) {
        self.borders.clear();
    }

    /// Cell margins/padding (`w:tcMar`).
    pub fn margins(&self) -> &CellMargins {
        &self.margins
    }

    /// Mutable cell margins.
    pub fn margins_mut(&mut self) -> &mut CellMargins {
        &mut self.margins
    }

    /// Set cell margins.
    pub fn set_margins(&mut self, margins: CellMargins) {
        self.margins = margins;
    }

    /// Clear all cell margins.
    pub fn clear_margins(&mut self) {
        self.margins.clear();
    }

    /// Unknown cell property children captured for roundtrip fidelity.
    pub(crate) fn unknown_property_children(&self) -> &[RawXmlNode] {
        self.unknown_property_children.as_slice()
    }

    /// Push an unknown cell property child node.
    pub(crate) fn push_unknown_property_child(&mut self, node: RawXmlNode) {
        self.unknown_property_children.push(node);
    }
}

/// Table alignment within the page (`w:jc`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableAlignment {
    /// Left-aligned (`w:jc w:val="start"` / `"left"`).
    Left,
    /// Center-aligned (`w:jc w:val="center"`).
    Center,
    /// Right-aligned (`w:jc w:val="end"` / `"right"`).
    Right,
}

/// Table width type (`w:tblW w:type`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableWidthType {
    /// Width in twips (`w:type="dxa"`).
    Dxa,
    /// Width as a percentage (`w:type="pct"`).
    Pct,
    /// Automatic width (`w:type="auto"`).
    Auto,
}

/// Table layout algorithm (`w:tblLayout`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableLayout {
    /// Fixed column widths (`w:tblLayout w:type="fixed"`).
    Fixed,
    /// Automatic column sizing (`w:tblLayout w:type="autofit"`).
    AutoFit,
}

/// A minimal table scaffold.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Table {
    rows: usize,
    columns: usize,
    cells: Vec<TableCell>,
    style_id: Option<String>,
    borders: TableBorders,
    /// Column widths from `w:tblGrid`, in twips.
    column_widths_twips: Vec<u32>,
    /// Per-row properties.
    row_properties: Vec<TableRowProperties>,
    /// Table alignment within the page (`w:jc`).
    alignment: Option<TableAlignment>,
    /// Table width value in twips (`w:tblW w:w`).
    width_twips: Option<u32>,
    /// Table width type (`w:tblW w:type`).
    width_type: Option<TableWidthType>,
    /// Table layout algorithm (`w:tblLayout`).
    autofit: Option<TableLayout>,
    /// Conditional formatting: first row (`w:tblLook w:firstRow`).
    first_row: bool,
    /// Conditional formatting: last row (`w:tblLook w:lastRow`).
    last_row: bool,
    /// Conditional formatting: first column (`w:tblLook w:firstColumn`).
    first_column: bool,
    /// Conditional formatting: last column (`w:tblLook w:lastColumn`).
    last_column: bool,
    /// Suppress horizontal banding (`w:tblLook w:noHBand`).
    no_h_band: bool,
    /// Suppress vertical banding (`w:tblLook w:noVBand`).
    no_v_band: bool,
    /// Unknown table property children captured for roundtrip fidelity.
    unknown_property_children: Vec<RawXmlNode>,
}

impl Table {
    /// Create a new table shape.
    pub fn new(rows: usize, columns: usize) -> Self {
        let cell_count = rows.checked_mul(columns).unwrap_or(0);
        let mut cells = Vec::new();
        cells.resize_with(cell_count, TableCell::new);

        let mut row_properties = Vec::new();
        row_properties.resize_with(rows, TableRowProperties::new);

        Self {
            rows,
            columns,
            cells,
            style_id: None,
            borders: TableBorders::new(),
            column_widths_twips: Vec::new(),
            row_properties,
            alignment: None,
            width_twips: None,
            width_type: None,
            autofit: None,
            first_row: false,
            last_row: false,
            first_column: false,
            last_column: false,
            no_h_band: false,
            no_v_band: false,
            unknown_property_children: Vec::new(),
        }
    }

    /// Number of rows.
    pub fn rows(&self) -> usize {
        self.rows
    }

    /// Number of columns.
    pub fn columns(&self) -> usize {
        self.columns
    }

    /// Whether the table has no rows or no columns.
    pub fn is_empty(&self) -> bool {
        self.rows == 0 || self.columns == 0
    }

    /// Table style identifier (`w:tblStyle`).
    pub fn style_id(&self) -> Option<&str> {
        self.style_id.as_deref()
    }

    /// Set table style identifier (`w:tblStyle`).
    pub fn set_style_id(&mut self, style_id: impl Into<String>) {
        self.style_id = normalize_optional_text(style_id.into());
    }

    /// Clear table style identifier.
    pub fn clear_style_id(&mut self) {
        self.style_id = None;
    }

    pub(crate) fn set_style_id_option(&mut self, style_id: Option<String>) {
        self.style_id = style_id.and_then(normalize_optional_text);
    }

    /// Get table borders (`w:tblBorders`).
    pub fn borders(&self) -> &TableBorders {
        &self.borders
    }

    /// Get mutable table borders (`w:tblBorders`).
    pub fn borders_mut(&mut self) -> &mut TableBorders {
        &mut self.borders
    }

    /// Replace table borders.
    pub fn set_borders(&mut self, borders: TableBorders) {
        self.borders = borders;
    }

    /// Clear all table borders.
    pub fn clear_borders(&mut self) {
        self.borders.clear();
    }

    /// Table alignment within the page (`w:jc`).
    pub fn alignment(&self) -> Option<TableAlignment> {
        self.alignment
    }

    /// Set table alignment within the page.
    pub fn set_alignment(&mut self, alignment: TableAlignment) {
        self.alignment = Some(alignment);
    }

    /// Clear table alignment.
    pub fn clear_alignment(&mut self) {
        self.alignment = None;
    }

    /// Table width in twips (`w:tblW w:w`).
    pub fn width_twips(&self) -> Option<u32> {
        self.width_twips
    }

    /// Set table width in twips.
    pub fn set_width_twips(&mut self, width: u32) {
        self.width_twips = Some(width);
    }

    /// Clear table width.
    pub fn clear_width_twips(&mut self) {
        self.width_twips = None;
    }

    /// Table width type (`w:tblW w:type`).
    pub fn width_type(&self) -> Option<TableWidthType> {
        self.width_type
    }

    /// Set table width type.
    pub fn set_width_type(&mut self, width_type: TableWidthType) {
        self.width_type = Some(width_type);
    }

    /// Table layout algorithm (`w:tblLayout`).
    pub fn layout(&self) -> Option<TableLayout> {
        self.autofit
    }

    /// Set table layout algorithm.
    pub fn set_layout(&mut self, layout: TableLayout) {
        self.autofit = Some(layout);
    }

    /// Clear table layout algorithm.
    pub fn clear_layout(&mut self) {
        self.autofit = None;
    }

    /// Whether conditional formatting applies to the first row (`w:tblLook w:firstRow`).
    pub fn first_row(&self) -> bool {
        self.first_row
    }

    /// Set conditional formatting for the first row.
    pub fn set_first_row(&mut self, value: bool) {
        self.first_row = value;
    }

    /// Whether conditional formatting applies to the last row (`w:tblLook w:lastRow`).
    pub fn last_row(&self) -> bool {
        self.last_row
    }

    /// Set conditional formatting for the last row.
    pub fn set_last_row(&mut self, value: bool) {
        self.last_row = value;
    }

    /// Whether conditional formatting applies to the first column (`w:tblLook w:firstColumn`).
    pub fn first_column(&self) -> bool {
        self.first_column
    }

    /// Set conditional formatting for the first column.
    pub fn set_first_column(&mut self, value: bool) {
        self.first_column = value;
    }

    /// Whether conditional formatting applies to the last column (`w:tblLook w:lastColumn`).
    pub fn last_column(&self) -> bool {
        self.last_column
    }

    /// Set conditional formatting for the last column.
    pub fn set_last_column(&mut self, value: bool) {
        self.last_column = value;
    }

    /// Whether horizontal banding is suppressed (`w:tblLook w:noHBand`).
    pub fn no_h_band(&self) -> bool {
        self.no_h_band
    }

    /// Set horizontal banding suppression.
    pub fn set_no_h_band(&mut self, value: bool) {
        self.no_h_band = value;
    }

    /// Whether vertical banding is suppressed (`w:tblLook w:noVBand`).
    pub fn no_v_band(&self) -> bool {
        self.no_v_band
    }

    /// Set vertical banding suppression.
    pub fn set_no_v_band(&mut self, value: bool) {
        self.no_v_band = value;
    }

    /// Get a cell by position.
    pub fn cell(&self, row: usize, column: usize) -> Option<&TableCell> {
        let index = self.cell_index(row, column)?;
        self.cells.get(index)
    }

    /// Get a mutable cell by position.
    pub fn cell_mut(&mut self, row: usize, column: usize) -> Option<&mut TableCell> {
        let index = self.cell_index(row, column)?;
        self.cells.get_mut(index)
    }

    /// Get a cell's text by position.
    pub fn cell_text(&self, row: usize, column: usize) -> Option<&str> {
        self.cell(row, column).map(TableCell::text)
    }

    /// Set a cell's text. Returns `false` if out of bounds.
    pub fn set_cell_text(&mut self, row: usize, column: usize, text: impl Into<String>) -> bool {
        let Some(cell) = self.cell_mut(row, column) else {
            return false;
        };
        cell.set_text(text);
        true
    }

    /// Merge horizontally adjacent cells in a row.
    pub fn merge_cells_horizontally(
        &mut self,
        row: usize,
        start_column: usize,
        span: usize,
    ) -> bool {
        if span < 2 || row >= self.rows || start_column >= self.columns {
            return false;
        }
        let Some(end_column_exclusive) = start_column.checked_add(span) else {
            return false;
        };
        if end_column_exclusive > self.columns {
            return false;
        }

        for column in start_column..end_column_exclusive {
            let _ = self.clear_horizontal_merge(row, column);
        }

        let Some(start_index) = self.cell_index(row, start_column) else {
            return false;
        };
        self.cells[start_index].horizontal_merge_continuation = false;
        self.cells[start_index].horizontal_span = span;

        for column in (start_column + 1)..end_column_exclusive {
            if let Some(index) = self.cell_index(row, column) {
                self.cells[index].horizontal_merge_continuation = true;
                self.cells[index].horizontal_span = 1;
                self.cells[index].text.clear();
            }
        }

        true
    }

    /// Clear a horizontal merge that includes this cell.
    pub fn clear_horizontal_merge(&mut self, row: usize, column: usize) -> bool {
        let Some(start_column) = self.find_horizontal_merge_start(row, column) else {
            return false;
        };
        self.clear_horizontal_merge_at_start(row, start_column)
    }

    fn clear_horizontal_merge_at_start(&mut self, row: usize, start_column: usize) -> bool {
        let Some(start_index) = self.cell_index(row, start_column) else {
            return false;
        };
        let span = self.cells[start_index].horizontal_span;
        if span <= 1 || self.cells[start_index].horizontal_merge_continuation {
            return false;
        }

        let end_column_exclusive = start_column.saturating_add(span).min(self.columns);
        for column in (start_column + 1)..end_column_exclusive {
            if let Some(index) = self.cell_index(row, column) {
                self.cells[index].horizontal_merge_continuation = false;
                self.cells[index].horizontal_span = 1;
            }
        }
        self.cells[start_index].horizontal_span = 1;
        true
    }

    fn find_horizontal_merge_start(&self, row: usize, column: usize) -> Option<usize> {
        let index = self.cell_index(row, column)?;
        let cell = self.cells.get(index)?;

        if !cell.horizontal_merge_continuation {
            return (cell.horizontal_span > 1).then_some(column);
        }

        for start_column in (0..column).rev() {
            let Some(start_index) = self.cell_index(row, start_column) else {
                continue;
            };
            let start_cell = &self.cells[start_index];
            if start_cell.horizontal_merge_continuation || start_cell.horizontal_span <= 1 {
                continue;
            }
            let covered_end = start_column.saturating_add(start_cell.horizontal_span);
            if covered_end > column {
                return Some(start_column);
            }
        }

        None
    }

    /// Column widths from `w:tblGrid`, in twips.
    pub fn column_widths_twips(&self) -> &[u32] {
        &self.column_widths_twips
    }

    /// Set column widths (populates `w:tblGrid`).
    pub fn set_column_widths_twips(&mut self, widths: Vec<u32>) {
        self.column_widths_twips = widths;
    }

    /// Clear column widths.
    pub fn clear_column_widths_twips(&mut self) {
        self.column_widths_twips.clear();
    }

    /// Get row properties for a given row.
    pub fn row_properties(&self, row: usize) -> Option<&TableRowProperties> {
        self.row_properties.get(row)
    }

    /// Get mutable row properties for a given row.
    pub fn row_properties_mut(&mut self, row: usize) -> Option<&mut TableRowProperties> {
        self.row_properties.get_mut(row)
    }

    /// Set row height in twips. Returns `false` if row out of bounds.
    pub fn set_row_height_twips(&mut self, row: usize, height: u32) -> bool {
        let Some(props) = self.row_properties.get_mut(row) else {
            return false;
        };
        props.set_height_twips(height);
        true
    }

    /// Set repeat header row flag. Returns `false` if row out of bounds.
    pub fn set_row_repeat_header(&mut self, row: usize, repeat: bool) -> bool {
        let Some(props) = self.row_properties.get_mut(row) else {
            return false;
        };
        props.set_repeat_header(repeat);
        true
    }

    pub(crate) fn set_horizontal_span(&mut self, row: usize, column: usize, span: usize) -> bool {
        if span <= 1 {
            return self.clear_horizontal_merge(row, column) || self.cell(row, column).is_some();
        }
        self.merge_cells_horizontally(row, column, span)
    }

    /// Append a new row at the end with default cells. Returns the index of the new row.
    pub fn add_row(&mut self) -> usize {
        let new_row_index = self.rows;
        self.insert_row(new_row_index);
        new_row_index
    }

    /// Insert a new row at the given index with default cells.
    ///
    /// If `index` is greater than or equal to the current row count, the row is
    /// appended at the end.
    pub fn insert_row(&mut self, index: usize) {
        let insert_at = index.min(self.rows);
        let flat_index = insert_at * self.columns;
        for _ in 0..self.columns {
            self.cells.insert(flat_index, TableCell::new());
        }
        self.row_properties
            .insert(insert_at, TableRowProperties::new());
        self.rows += 1;
    }

    /// Remove the row at the given index.
    ///
    /// Returns `false` if the index is out of bounds.
    pub fn remove_row(&mut self, index: usize) -> bool {
        if index >= self.rows {
            return false;
        }
        let flat_start = index * self.columns;
        self.cells.drain(flat_start..flat_start + self.columns);
        self.row_properties.remove(index);
        self.rows -= 1;
        true
    }

    /// Append a new column at the end with default cells. Returns the index of the new column.
    pub fn add_column(&mut self) -> usize {
        let new_col_index = self.columns;
        self.insert_column(new_col_index);
        new_col_index
    }

    /// Insert a new column at the given index with default cells.
    ///
    /// If `index` is greater than or equal to the current column count, the
    /// column is appended at the end.
    pub fn insert_column(&mut self, index: usize) {
        let insert_at = index.min(self.columns);
        // Insert one cell per row, working from the last row backward so
        // earlier insertions don't shift the indices of later rows.
        for row in (0..self.rows).rev() {
            let flat_index = row * self.columns + insert_at;
            self.cells.insert(flat_index, TableCell::new());
        }
        self.columns += 1;
    }

    /// Remove the column at the given index.
    ///
    /// Returns `false` if the index is out of bounds.
    pub fn remove_column(&mut self, index: usize) -> bool {
        if index >= self.columns {
            return false;
        }
        // Remove one cell per row, working from the last row backward so
        // earlier removals don't shift the indices of later rows.
        for row in (0..self.rows).rev() {
            let flat_index = row * self.columns + index;
            self.cells.remove(flat_index);
        }
        self.columns -= 1;
        true
    }

    fn cell_index(&self, row: usize, column: usize) -> Option<usize> {
        if row >= self.rows || column >= self.columns {
            return None;
        }

        row.checked_mul(self.columns)?.checked_add(column)
    }

    /// Unknown table property children captured for roundtrip fidelity.
    pub(crate) fn unknown_property_children(&self) -> &[RawXmlNode] {
        self.unknown_property_children.as_slice()
    }

    /// Push an unknown table property child node.
    pub(crate) fn push_unknown_property_child(&mut self, node: RawXmlNode) {
        self.unknown_property_children.push(node);
    }
}

fn normalize_optional_text(value: String) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn normalize_color_value(value: &str) -> Option<String> {
    let trimmed = value.trim();
    let normalized = trimmed.trim_start_matches('#').to_ascii_uppercase();
    if normalized.is_empty() {
        None
    } else {
        Some(normalized)
    }
}

#[cfg(test)]
mod tests {
    use super::{Table, TableBorder, VerticalAlignment, VerticalMerge};

    #[test]
    fn bounds_checked_cell_access() {
        let table = Table::new(2, 2);

        assert!(table.cell(0, 0).is_some());
        assert!(table.cell(2, 0).is_none());
        assert!(table.cell(0, 2).is_none());
    }

    #[test]
    fn set_and_get_cell_text() {
        let mut table = Table::new(2, 2);

        assert!(table.set_cell_text(1, 1, "Cell text"));
        assert_eq!(table.cell_text(1, 1), Some("Cell text"));
        assert!(!table.set_cell_text(2, 0, "Out of bounds"));
        assert_eq!(table.cell_text(2, 0), None);
    }

    #[test]
    fn horizontal_merge_flags_start_and_continuation_cells() {
        let mut table = Table::new(1, 4);
        assert!(table.set_cell_text(0, 0, "A"));
        assert!(table.set_cell_text(0, 1, "B"));
        assert!(table.set_cell_text(0, 2, "C"));

        assert!(table.merge_cells_horizontally(0, 1, 2));
        assert_eq!(table.cell(0, 1).map(|cell| cell.horizontal_span()), Some(2));
        assert_eq!(
            table
                .cell(0, 2)
                .map(|cell| cell.is_horizontal_merge_continuation()),
            Some(true)
        );
        assert_eq!(table.cell_text(0, 2), Some(""));

        assert!(table.clear_horizontal_merge(0, 2));
        assert_eq!(table.cell(0, 1).map(|cell| cell.horizontal_span()), Some(1));
        assert_eq!(
            table
                .cell(0, 2)
                .map(|cell| cell.is_horizontal_merge_continuation()),
            Some(false)
        );
    }

    #[test]
    fn table_borders_can_be_set() {
        let mut table = Table::new(1, 1);
        let mut top = TableBorder::new("single");
        top.set_size_eighth_points(8);
        top.set_color("#00aa11");
        table.borders_mut().set_top(top);

        assert_eq!(
            table.borders().top().and_then(TableBorder::line_type),
            Some("single")
        );
        assert_eq!(
            table
                .borders()
                .top()
                .and_then(TableBorder::size_eighth_points),
            Some(8)
        );
        assert_eq!(
            table.borders().top().and_then(TableBorder::color),
            Some("00AA11")
        );
    }

    #[test]
    fn table_style_id_can_be_set_and_cleared() {
        let mut table = Table::new(1, 1);

        table.set_style_id(" TableGrid ");
        assert_eq!(table.style_id(), Some("TableGrid"));

        table.clear_style_id();
        assert_eq!(table.style_id(), None);
    }

    #[test]
    fn vertical_merge_can_be_set_and_cleared() {
        let mut table = Table::new(2, 1);
        let cell = table.cell_mut(0, 0).expect("cell must exist");

        assert_eq!(cell.vertical_merge(), None);

        cell.set_vertical_merge(VerticalMerge::Restart);
        assert_eq!(cell.vertical_merge(), Some(VerticalMerge::Restart));

        let cell = table.cell_mut(1, 0).expect("cell must exist");
        cell.set_vertical_merge(VerticalMerge::Continue);
        assert_eq!(cell.vertical_merge(), Some(VerticalMerge::Continue));

        cell.clear_vertical_merge();
        assert_eq!(cell.vertical_merge(), None);
    }

    #[test]
    fn shading_color_can_be_set_and_cleared() {
        let mut table = Table::new(1, 1);
        let cell = table.cell_mut(0, 0).expect("cell must exist");

        assert_eq!(cell.shading_color(), None);

        cell.set_shading_color("#ff0000");
        assert_eq!(cell.shading_color(), Some("FF0000"));

        cell.set_shading_color("aabb11");
        assert_eq!(cell.shading_color(), Some("AABB11"));

        cell.clear_shading_color();
        assert_eq!(cell.shading_color(), None);
    }

    #[test]
    fn vertical_alignment_can_be_set_and_cleared() {
        let mut table = Table::new(1, 1);
        let cell = table.cell_mut(0, 0).expect("cell must exist");

        assert_eq!(cell.vertical_alignment(), None);

        cell.set_vertical_alignment(VerticalAlignment::Center);
        assert_eq!(cell.vertical_alignment(), Some(VerticalAlignment::Center));

        cell.set_vertical_alignment(VerticalAlignment::Bottom);
        assert_eq!(cell.vertical_alignment(), Some(VerticalAlignment::Bottom));

        cell.set_vertical_alignment(VerticalAlignment::Top);
        assert_eq!(cell.vertical_alignment(), Some(VerticalAlignment::Top));

        cell.clear_vertical_alignment();
        assert_eq!(cell.vertical_alignment(), None);
    }

    #[test]
    fn cell_width_twips_can_be_set_and_cleared() {
        let mut table = Table::new(1, 1);
        let cell = table.cell_mut(0, 0).expect("cell must exist");

        assert_eq!(cell.cell_width_twips(), None);

        cell.set_cell_width_twips(2400);
        assert_eq!(cell.cell_width_twips(), Some(2400));

        cell.clear_cell_width_twips();
        assert_eq!(cell.cell_width_twips(), None);
    }

    #[test]
    fn cell_borders_can_be_set_and_cleared() {
        let mut table = Table::new(1, 1);
        let cell = table.cell_mut(0, 0).expect("cell must exist");
        assert!(cell.borders().is_empty());

        let mut top = TableBorder::new("single");
        top.set_size_eighth_points(4);
        top.set_color("000000");
        cell.borders_mut().set_top(top);

        assert_eq!(
            cell.borders().top().and_then(TableBorder::line_type),
            Some("single")
        );
        assert_eq!(
            cell.borders()
                .top()
                .and_then(TableBorder::size_eighth_points),
            Some(4)
        );

        cell.clear_borders();
        assert!(cell.borders().is_empty());
    }

    #[test]
    fn cell_margins_can_be_set_and_cleared() {
        let mut table = Table::new(1, 1);
        let cell = table.cell_mut(0, 0).expect("cell must exist");
        assert!(cell.margins().is_empty());

        cell.margins_mut().set_top_twips(100);
        cell.margins_mut().set_left_twips(200);
        cell.margins_mut().set_bottom_twips(100);
        cell.margins_mut().set_right_twips(200);

        assert_eq!(cell.margins().top_twips(), Some(100));
        assert_eq!(cell.margins().left_twips(), Some(200));
        assert_eq!(cell.margins().bottom_twips(), Some(100));
        assert_eq!(cell.margins().right_twips(), Some(200));

        cell.clear_margins();
        assert!(cell.margins().is_empty());
    }

    #[test]
    fn column_widths_can_be_set_and_cleared() {
        let mut table = Table::new(2, 3);
        assert!(table.column_widths_twips().is_empty());

        table.set_column_widths_twips(vec![2400, 3600, 2400]);
        assert_eq!(table.column_widths_twips(), &[2400, 3600, 2400]);

        table.clear_column_widths_twips();
        assert!(table.column_widths_twips().is_empty());
    }

    #[test]
    fn row_properties_can_be_set() {
        let mut table = Table::new(3, 2);
        assert!(table.set_row_height_twips(0, 720));
        assert!(table.set_row_repeat_header(0, true));

        let props = table.row_properties(0).expect("row 0 props must exist");
        assert_eq!(props.height_twips(), Some(720));
        assert!(props.repeat_header());

        let props = table.row_properties(1).expect("row 1 props must exist");
        assert_eq!(props.height_twips(), None);
        assert!(!props.repeat_header());

        assert!(!table.set_row_height_twips(5, 100));
    }
}
