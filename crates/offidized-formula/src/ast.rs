use crate::reference::{CellRef, RangeRef};
use crate::value::ScalarValue;

/// A parsed formula expression.
#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    /// A literal scalar value.
    Literal(ScalarValue),
    /// A reference to a single cell.
    CellRef(CellRef),
    /// A reference to a range of cells.
    RangeRef(RangeRef),
    /// A unary operation.
    Unary { op: UnaryOp, expr: Box<Expr> },
    /// A binary operation.
    Binary {
        op: BinaryOp,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    /// A function call.
    Function { name: String, args: Vec<Expr> },
    /// A defined name reference.
    Name(String),
}

/// Unary operators.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOp {
    /// Unary `+` (identity).
    Plus,
    /// Unary `-` (negation).
    Negate,
    /// Postfix `%` (divide by 100).
    Percent,
}

/// Binary operators in order of increasing precedence group.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOp {
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
    /// `&` (string concatenation)
    Concat,
    /// `+`
    Add,
    /// `-`
    Sub,
    /// `*`
    Mul,
    /// `/`
    Div,
    /// `^` (exponentiation, right-associative)
    Pow,
    /// `:` (range operator)
    Range,
}
