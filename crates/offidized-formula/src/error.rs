/// Errors that can occur during formula parsing and evaluation.
#[derive(Debug, thiserror::Error)]
pub enum FormulaError {
    /// The formula string could not be parsed.
    #[error("parse error: {0}")]
    Parse(String),

    /// A lexer error occurred while tokenizing the formula.
    #[error("lexer error: {0}")]
    Lex(String),

    /// An unknown function was called.
    #[error("unknown function: {0}")]
    UnknownFunction(String),

    /// Wrong number of arguments passed to a function.
    #[error("function {name} expects {expected} arguments, got {got}")]
    ArgumentCount {
        name: String,
        expected: String,
        got: usize,
    },

    /// A reference could not be resolved.
    #[error("unresolvable reference: {0}")]
    Reference(String),

    /// A named range or defined name could not be resolved.
    #[error("unknown name: {0}")]
    UnknownName(String),
}

/// A specialized `Result` type for formula operations.
pub type Result<T> = std::result::Result<T, FormulaError>;
