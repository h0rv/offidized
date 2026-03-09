/// Print settings types for worksheet printing (ECMA-376 18.3.1.40, 18.3.1.58, 18.3.1.73).
///
/// These complement the existing `PageSetup` and `PageMargins` types with header/footer
/// content, print area designations, and page break definitions.
/// Header and footer content for printed pages.
///
/// OOXML supports different headers/footers for odd pages, even pages, and the
/// first page. The `different_odd_even` and `different_first` flags control whether
/// those alternates are active.
///
/// Header/footer strings use OOXML formatting codes (e.g. `&L`, `&C`, `&R` for
/// left/center/right sections, `&P` for page number, `&D` for date, etc.).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PrintHeaderFooter {
    odd_header: Option<String>,
    odd_footer: Option<String>,
    even_header: Option<String>,
    even_footer: Option<String>,
    first_header: Option<String>,
    first_footer: Option<String>,
    different_odd_even: bool,
    different_first: bool,
}

impl PrintHeaderFooter {
    /// Creates a new empty header/footer configuration.
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns the odd-page header string.
    pub fn odd_header(&self) -> Option<&str> {
        self.odd_header.as_deref()
    }

    /// Sets the odd-page header string.
    pub fn set_odd_header(&mut self, header: impl Into<String>) -> &mut Self {
        self.odd_header = Some(header.into());
        self
    }

    /// Clears the odd-page header.
    pub fn clear_odd_header(&mut self) -> &mut Self {
        self.odd_header = None;
        self
    }

    /// Returns the odd-page footer string.
    pub fn odd_footer(&self) -> Option<&str> {
        self.odd_footer.as_deref()
    }

    /// Sets the odd-page footer string.
    pub fn set_odd_footer(&mut self, footer: impl Into<String>) -> &mut Self {
        self.odd_footer = Some(footer.into());
        self
    }

    /// Clears the odd-page footer.
    pub fn clear_odd_footer(&mut self) -> &mut Self {
        self.odd_footer = None;
        self
    }

    /// Returns the even-page header string.
    pub fn even_header(&self) -> Option<&str> {
        self.even_header.as_deref()
    }

    /// Sets the even-page header string.
    pub fn set_even_header(&mut self, header: impl Into<String>) -> &mut Self {
        self.even_header = Some(header.into());
        self
    }

    /// Clears the even-page header.
    pub fn clear_even_header(&mut self) -> &mut Self {
        self.even_header = None;
        self
    }

    /// Returns the even-page footer string.
    pub fn even_footer(&self) -> Option<&str> {
        self.even_footer.as_deref()
    }

    /// Sets the even-page footer string.
    pub fn set_even_footer(&mut self, footer: impl Into<String>) -> &mut Self {
        self.even_footer = Some(footer.into());
        self
    }

    /// Clears the even-page footer.
    pub fn clear_even_footer(&mut self) -> &mut Self {
        self.even_footer = None;
        self
    }

    /// Returns the first-page header string.
    pub fn first_header(&self) -> Option<&str> {
        self.first_header.as_deref()
    }

    /// Sets the first-page header string.
    pub fn set_first_header(&mut self, header: impl Into<String>) -> &mut Self {
        self.first_header = Some(header.into());
        self
    }

    /// Clears the first-page header.
    pub fn clear_first_header(&mut self) -> &mut Self {
        self.first_header = None;
        self
    }

    /// Returns the first-page footer string.
    pub fn first_footer(&self) -> Option<&str> {
        self.first_footer.as_deref()
    }

    /// Sets the first-page footer string.
    pub fn set_first_footer(&mut self, footer: impl Into<String>) -> &mut Self {
        self.first_footer = Some(footer.into());
        self
    }

    /// Clears the first-page footer.
    pub fn clear_first_footer(&mut self) -> &mut Self {
        self.first_footer = None;
        self
    }

    /// Returns whether different odd/even headers/footers are enabled.
    pub fn different_odd_even(&self) -> bool {
        self.different_odd_even
    }

    /// Sets whether different odd/even headers/footers are enabled.
    pub fn set_different_odd_even(&mut self, value: bool) -> &mut Self {
        self.different_odd_even = value;
        self
    }

    /// Returns whether a different first-page header/footer is enabled.
    pub fn different_first(&self) -> bool {
        self.different_first
    }

    /// Sets whether a different first-page header/footer is enabled.
    pub fn set_different_first(&mut self, value: bool) -> &mut Self {
        self.different_first = value;
        self
    }

    /// Returns `true` if any header/footer content or flags are set.
    pub(crate) fn has_metadata(&self) -> bool {
        self.odd_header.is_some()
            || self.odd_footer.is_some()
            || self.even_header.is_some()
            || self.even_footer.is_some()
            || self.first_header.is_some()
            || self.first_footer.is_some()
            || self.different_odd_even
            || self.different_first
    }
}

/// A print area defined as a cell range string (e.g. `"A1:H50"`).
///
/// This is a newtype wrapper around `String` for type safety. The print area
/// specifies which cells should be included when the worksheet is printed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrintArea(String);

impl PrintArea {
    /// Creates a new print area from a cell range string.
    pub fn new(range: impl Into<String>) -> Self {
        Self(range.into())
    }

    /// Returns the cell range string.
    pub fn range(&self) -> &str {
        &self.0
    }

    /// Sets the cell range string.
    pub fn set_range(&mut self, range: impl Into<String>) -> &mut Self {
        self.0 = range.into();
        self
    }
}

impl std::fmt::Display for PrintArea {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

impl From<&str> for PrintArea {
    fn from(s: &str) -> Self {
        Self(s.to_string())
    }
}

impl From<String> for PrintArea {
    fn from(s: String) -> Self {
        Self(s)
    }
}

/// A single page break at a row or column index.
///
/// In OOXML, page breaks are recorded with an `id` (the 1-based row or column index
/// where the break occurs) and a `manual` flag indicating whether the user inserted
/// the break explicitly.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageBreak {
    id: u32,
    manual: bool,
}

impl PageBreak {
    /// Creates a new manual page break at the given row or column index.
    pub fn new(id: u32) -> Self {
        Self { id, manual: true }
    }

    /// Creates a new page break with explicit manual flag.
    pub fn with_manual(id: u32, manual: bool) -> Self {
        Self { id, manual }
    }

    /// Returns the 1-based row or column index of this break.
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Sets the row or column index.
    pub fn set_id(&mut self, id: u32) -> &mut Self {
        self.id = id;
        self
    }

    /// Returns whether this is a manual (user-inserted) break.
    pub fn manual(&self) -> bool {
        self.manual
    }

    /// Sets the manual flag.
    pub fn set_manual(&mut self, manual: bool) -> &mut Self {
        self.manual = manual;
        self
    }
}

/// Collection of row and column page breaks for a worksheet.
///
/// OOXML stores row breaks (`<rowBreaks>`) and column breaks (`<colBreaks>`)
/// separately. This struct holds both collections.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PageBreaks {
    row_breaks: Vec<PageBreak>,
    col_breaks: Vec<PageBreak>,
}

impl PageBreaks {
    /// Creates a new empty page breaks collection.
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns the row breaks.
    pub fn row_breaks(&self) -> &[PageBreak] {
        &self.row_breaks
    }

    /// Returns mutable access to the row breaks.
    pub fn row_breaks_mut(&mut self) -> &mut Vec<PageBreak> {
        &mut self.row_breaks
    }

    /// Adds a row break at the given 1-based row index.
    pub fn add_row_break(&mut self, row: u32) -> &mut Self {
        self.row_breaks.push(PageBreak::new(row));
        self
    }

    /// Adds a row break with explicit manual flag.
    pub fn add_row_break_with_manual(&mut self, row: u32, manual: bool) -> &mut Self {
        self.row_breaks.push(PageBreak::with_manual(row, manual));
        self
    }

    /// Clears all row breaks.
    pub fn clear_row_breaks(&mut self) -> &mut Self {
        self.row_breaks.clear();
        self
    }

    /// Returns the column breaks.
    pub fn col_breaks(&self) -> &[PageBreak] {
        &self.col_breaks
    }

    /// Returns mutable access to the column breaks.
    pub fn col_breaks_mut(&mut self) -> &mut Vec<PageBreak> {
        &mut self.col_breaks
    }

    /// Adds a column break at the given 1-based column index.
    pub fn add_col_break(&mut self, col: u32) -> &mut Self {
        self.col_breaks.push(PageBreak::new(col));
        self
    }

    /// Adds a column break with explicit manual flag.
    pub fn add_col_break_with_manual(&mut self, col: u32, manual: bool) -> &mut Self {
        self.col_breaks.push(PageBreak::with_manual(col, manual));
        self
    }

    /// Clears all column breaks.
    pub fn clear_col_breaks(&mut self) -> &mut Self {
        self.col_breaks.clear();
        self
    }

    /// Clears all row and column breaks.
    pub fn clear_all(&mut self) -> &mut Self {
        self.row_breaks.clear();
        self.col_breaks.clear();
        self
    }

    /// Returns `true` if there are any row or column breaks.
    pub fn has_breaks(&self) -> bool {
        !self.row_breaks.is_empty() || !self.col_breaks.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn print_header_footer_defaults() {
        let hf = PrintHeaderFooter::new();
        assert_eq!(hf.odd_header(), None);
        assert_eq!(hf.odd_footer(), None);
        assert_eq!(hf.even_header(), None);
        assert_eq!(hf.even_footer(), None);
        assert_eq!(hf.first_header(), None);
        assert_eq!(hf.first_footer(), None);
        assert!(!hf.different_odd_even());
        assert!(!hf.different_first());
        assert!(!hf.has_metadata());
    }

    #[test]
    fn print_header_footer_setters() {
        let mut hf = PrintHeaderFooter::new();
        hf.set_odd_header("&CPage &P of &N")
            .set_odd_footer("&LConfidential&R&D")
            .set_different_odd_even(true)
            .set_even_header("&CEven Page &P")
            .set_even_footer("&LEven Footer")
            .set_different_first(true)
            .set_first_header("&CFirst Page Header")
            .set_first_footer("&CFirst Page Footer");

        assert_eq!(hf.odd_header(), Some("&CPage &P of &N"));
        assert_eq!(hf.odd_footer(), Some("&LConfidential&R&D"));
        assert_eq!(hf.even_header(), Some("&CEven Page &P"));
        assert_eq!(hf.even_footer(), Some("&LEven Footer"));
        assert_eq!(hf.first_header(), Some("&CFirst Page Header"));
        assert_eq!(hf.first_footer(), Some("&CFirst Page Footer"));
        assert!(hf.different_odd_even());
        assert!(hf.different_first());
        assert!(hf.has_metadata());
    }

    #[test]
    fn print_header_footer_clear() {
        let mut hf = PrintHeaderFooter::new();
        hf.set_odd_header("Header")
            .set_odd_footer("Footer")
            .set_even_header("Even Header")
            .set_even_footer("Even Footer")
            .set_first_header("First Header")
            .set_first_footer("First Footer");

        hf.clear_odd_header()
            .clear_odd_footer()
            .clear_even_header()
            .clear_even_footer()
            .clear_first_header()
            .clear_first_footer()
            .set_different_odd_even(false)
            .set_different_first(false);

        assert!(!hf.has_metadata());
    }

    #[test]
    fn print_area_construction() {
        let area = PrintArea::new("A1:H50");
        assert_eq!(area.range(), "A1:H50");
        assert_eq!(area.to_string(), "A1:H50");

        let area2: PrintArea = "B2:G100".into();
        assert_eq!(area2.range(), "B2:G100");

        let area3: PrintArea = String::from("Sheet1!A1:Z100").into();
        assert_eq!(area3.range(), "Sheet1!A1:Z100");
    }

    #[test]
    fn print_area_set_range() {
        let mut area = PrintArea::new("A1:B10");
        area.set_range("C1:D20");
        assert_eq!(area.range(), "C1:D20");
    }

    #[test]
    fn page_break_construction() {
        let brk = PageBreak::new(10);
        assert_eq!(brk.id(), 10);
        assert!(brk.manual());

        let brk2 = PageBreak::with_manual(5, false);
        assert_eq!(brk2.id(), 5);
        assert!(!brk2.manual());
    }

    #[test]
    fn page_break_setters() {
        let mut brk = PageBreak::new(10);
        brk.set_id(20).set_manual(false);
        assert_eq!(brk.id(), 20);
        assert!(!brk.manual());
    }

    #[test]
    fn page_breaks_row_and_col() {
        let mut breaks = PageBreaks::new();
        assert!(!breaks.has_breaks());

        breaks.add_row_break(10).add_row_break(20);
        breaks.add_col_break(5);
        assert!(breaks.has_breaks());
        assert_eq!(breaks.row_breaks().len(), 2);
        assert_eq!(breaks.col_breaks().len(), 1);
        assert_eq!(breaks.row_breaks()[0].id(), 10);
        assert_eq!(breaks.row_breaks()[1].id(), 20);
        assert_eq!(breaks.col_breaks()[0].id(), 5);
    }

    #[test]
    fn page_breaks_with_manual_flag() {
        let mut breaks = PageBreaks::new();
        breaks.add_row_break_with_manual(15, false);
        breaks.add_col_break_with_manual(3, true);

        assert!(!breaks.row_breaks()[0].manual());
        assert!(breaks.col_breaks()[0].manual());
    }

    #[test]
    fn page_breaks_clear() {
        let mut breaks = PageBreaks::new();
        breaks.add_row_break(10).add_col_break(5);

        breaks.clear_row_breaks();
        assert!(breaks.row_breaks().is_empty());
        assert!(!breaks.col_breaks().is_empty());

        breaks.add_row_break(20);
        breaks.clear_all();
        assert!(!breaks.has_breaks());
    }

    #[test]
    fn page_breaks_mut_access() {
        let mut breaks = PageBreaks::new();
        breaks.add_row_break(10);
        breaks.row_breaks_mut()[0].set_id(15);
        assert_eq!(breaks.row_breaks()[0].id(), 15);

        breaks.add_col_break(3);
        breaks.col_breaks_mut()[0].set_manual(false);
        assert!(!breaks.col_breaks()[0].manual());
    }
}
