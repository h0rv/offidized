use crate::value::XlError;

/// A lexical token produced by the formula scanner.
#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    /// A numeric literal (e.g. `42`, `3.14`, `1.5E-3`).
    Number(f64),
    /// A string literal (e.g. `"hello"`).
    StringLiteral(String),
    /// A boolean literal (`TRUE` or `FALSE`).
    Bool(bool),
    /// An error literal (e.g. `#DIV/0!`).
    Error(XlError),
    /// A cell reference like `A1`, `$B$3`, with optional sheet prefix already stripped.
    CellReference {
        sheet: Option<String>,
        col_letters: String,
        row: u32,
        col_absolute: bool,
        row_absolute: bool,
    },
    /// A function name or identifier (e.g. `SUM`, `myName`).
    Ident(String),
    /// `+`
    Plus,
    /// `-`
    Minus,
    /// `*`
    Star,
    /// `/`
    Slash,
    /// `^`
    Caret,
    /// `%`
    Percent,
    /// `&` (string concatenation)
    Ampersand,
    /// `=`
    Eq,
    /// `<>`
    Ne,
    /// `<`
    Lt,
    /// `<=`
    Le,
    /// `>`
    Gt,
    /// `>=`
    Ge,
    /// `:`
    Colon,
    /// `,`
    Comma,
    /// `(`
    LParen,
    /// `)`
    RParen,
    /// End of formula.
    Eof,
}

impl Token {
    /// Returns true if this token represents a cell reference.
    pub fn is_cell_ref(&self) -> bool {
        matches!(self, Self::CellReference { .. })
    }
}
