use crate::value::ScalarValue;

/// Trait for providing cell data to the formula evaluator.
///
/// Consumers implement this trait to bridge their data model (e.g. a worksheet)
/// to the formula engine. The evaluator calls these methods to resolve cell
/// references, defined names, and sheet lookups during evaluation.
pub trait CellDataProvider {
    /// Returns the value of a cell.
    ///
    /// `sheet` is `None` for the current sheet (the sheet the formula lives in).
    /// `row` and `col` are 1-based indices.
    fn cell_value(&self, sheet: Option<&str>, row: u32, col: u32) -> ScalarValue;

    /// Returns the formula text of a cell, if any.
    ///
    /// This is used for dependency analysis and future support for recursive
    /// evaluation. The default implementation returns `None` (no formula).
    fn cell_formula(&self, _sheet: Option<&str>, _row: u32, _col: u32) -> Option<String> {
        None
    }

    /// Resolves a defined name to a formula/value string.
    ///
    /// For example, resolving `"TaxRate"` might return `"0.08"` or `"Sheet1!$A$1"`.
    /// The default implementation returns `None` (name not found).
    fn resolve_name(&self, _name: &str) -> Option<String> {
        None
    }

    /// Returns the name of the sheet at the given 0-based index.
    ///
    /// Used for 3-D references and sheet enumeration. The default implementation
    /// returns `None`.
    fn sheet_name(&self, _index: usize) -> Option<String> {
        None
    }

    /// Returns the 0-based index of a sheet by name.
    ///
    /// Used for 3-D references. The default implementation returns `None`.
    fn sheet_index(&self, _name: &str) -> Option<usize> {
        None
    }

    /// Returns cell information for the CELL function.
    ///
    /// `info_type` is one of: "address", "col", "row", "contents", "type", "format", "width".
    /// The default implementation returns `None`.
    fn cell_info(
        &self,
        _sheet: Option<&str>,
        _row: u32,
        _col: u32,
        _info_type: &str,
    ) -> Option<String> {
        None
    }

    /// Returns workbook/environment information for the INFO function.
    ///
    /// `info_type` is one of: "numfile", "recalc", "release", "system", "osversion", "directory", "origin".
    /// The default implementation returns `None`.
    fn workbook_info(&self, _info_type: &str) -> Option<String> {
        None
    }
}

/// Evaluation context passed to the formula evaluator.
///
/// Contains the data provider, the location of the formula being evaluated,
/// and configuration flags.
pub struct EvalContext<'a> {
    /// The cell data provider that supplies cell values and metadata.
    pub provider: &'a dyn CellDataProvider,
    /// The name of the sheet the formula is on, or `None` for a global context.
    pub current_sheet: Option<String>,
    /// 1-based row of the cell containing the formula.
    pub formula_row: u32,
    /// 1-based column of the cell containing the formula.
    pub formula_col: u32,
    /// Whether to use the 1904 date system (default: false, uses 1900).
    pub use_1904_date_system: bool,
}

impl<'a> EvalContext<'a> {
    /// Creates a new evaluation context.
    ///
    /// # Arguments
    ///
    /// * `provider` - The cell data provider
    /// * `current_sheet` - The name of the sheet the formula is on (or `None`)
    /// * `row` - 1-based row of the cell containing the formula
    /// * `col` - 1-based column of the cell containing the formula
    pub fn new(
        provider: &'a dyn CellDataProvider,
        current_sheet: Option<impl Into<String>>,
        row: u32,
        col: u32,
    ) -> Self {
        Self {
            provider,
            current_sheet: current_sheet.map(Into::into),
            formula_row: row,
            formula_col: col,
            use_1904_date_system: false,
        }
    }

    /// Creates a stub evaluation context for testing.
    ///
    /// The stub provider returns `Blank` for all cell values and `None` for all
    /// other queries. This is useful for testing functions that don't need
    /// cell data.
    #[cfg(test)]
    pub fn stub() -> Self {
        struct StubProvider;
        impl CellDataProvider for StubProvider {
            fn cell_value(&self, _sheet: Option<&str>, _row: u32, _col: u32) -> ScalarValue {
                ScalarValue::Blank
            }
        }
        static STUB: StubProvider = StubProvider;
        Self::new(&STUB, None::<String>, 1, 1)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    struct DummyProvider;
    impl CellDataProvider for DummyProvider {
        fn cell_value(&self, _sheet: Option<&str>, _row: u32, _col: u32) -> ScalarValue {
            ScalarValue::Blank
        }
    }

    #[test]
    fn eval_context_new_with_sheet() {
        let provider = DummyProvider;
        let ctx = EvalContext::new(&provider, Some("Sheet1"), 5, 3);
        assert_eq!(ctx.current_sheet, Some("Sheet1".to_string()));
        assert_eq!(ctx.formula_row, 5);
        assert_eq!(ctx.formula_col, 3);
        assert!(!ctx.use_1904_date_system);
    }

    #[test]
    fn eval_context_new_no_sheet() {
        let provider = DummyProvider;
        let ctx = EvalContext::new(&provider, None::<String>, 1, 1);
        assert_eq!(ctx.current_sheet, None);
    }

    #[test]
    fn default_trait_methods_return_none() {
        let provider = DummyProvider;
        assert_eq!(provider.cell_formula(None, 1, 1), None);
        assert_eq!(provider.resolve_name("anything"), None);
        assert_eq!(provider.sheet_name(0), None);
        assert_eq!(provider.sheet_index("Sheet1"), None);
    }
}
