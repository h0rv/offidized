use std::fmt;

/// Excel error values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlError {
    /// `#NULL!` — intersection of two ranges that don't intersect.
    Null,
    /// `#DIV/0!` — division by zero.
    Div0,
    /// `#VALUE!` — wrong type of argument or operand.
    Value,
    /// `#REF!` — invalid cell reference.
    Ref,
    /// `#NAME?` — unrecognised formula name.
    Name,
    /// `#NUM!` — invalid numeric value.
    Num,
    /// `#N/A` — value not available.
    Na,
}

impl fmt::Display for XlError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Null => write!(f, "#NULL!"),
            Self::Div0 => write!(f, "#DIV/0!"),
            Self::Value => write!(f, "#VALUE!"),
            Self::Ref => write!(f, "#REF!"),
            Self::Name => write!(f, "#NAME?"),
            Self::Num => write!(f, "#NUM!"),
            Self::Na => write!(f, "#N/A"),
        }
    }
}

impl XlError {
    /// Parses an error literal string into an `XlError`.
    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "#NULL!" => Some(Self::Null),
            "#DIV/0!" => Some(Self::Div0),
            "#VALUE!" => Some(Self::Value),
            "#REF!" => Some(Self::Ref),
            "#NAME?" => Some(Self::Name),
            "#NUM!" => Some(Self::Num),
            "#N/A" => Some(Self::Na),
            _ => None,
        }
    }
}

/// A single scalar value in the formula engine.
#[derive(Debug, Clone, PartialEq)]
pub enum ScalarValue {
    /// No value (empty cell).
    Blank,
    /// Boolean value.
    Bool(bool),
    /// Numeric value (all Excel numbers are f64).
    Number(f64),
    /// Text string.
    Text(String),
    /// An Excel error.
    Error(XlError),
}

impl fmt::Display for ScalarValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Blank => write!(f, ""),
            Self::Bool(b) => {
                if *b {
                    write!(f, "TRUE")
                } else {
                    write!(f, "FALSE")
                }
            }
            Self::Number(n) => write!(f, "{n}"),
            Self::Text(s) => write!(f, "{s}"),
            Self::Error(e) => write!(f, "{e}"),
        }
    }
}

impl ScalarValue {
    /// Coerces this value to a number, following Excel semantics.
    ///
    /// - Blank → 0.0
    /// - Bool: TRUE → 1.0, FALSE → 0.0
    /// - Number → as-is
    /// - Text → parse as f64 or `#VALUE!`
    /// - Error → propagated
    pub fn to_number(&self) -> ScalarValue {
        match self {
            Self::Blank => Self::Number(0.0),
            Self::Bool(b) => Self::Number(if *b { 1.0 } else { 0.0 }),
            Self::Number(_) => self.clone(),
            Self::Text(s) => {
                if s.is_empty() {
                    return Self::Number(0.0);
                }
                match s.trim().parse::<f64>() {
                    Ok(n) => Self::Number(n),
                    Err(_) => Self::Error(XlError::Value),
                }
            }
            Self::Error(_) => self.clone(),
        }
    }

    /// Coerces this value to text, following Excel semantics.
    ///
    /// - Blank → ""
    /// - Bool → "TRUE" / "FALSE"
    /// - Number → decimal representation
    /// - Text → as-is
    /// - Error → propagated
    pub fn to_text(&self) -> ScalarValue {
        match self {
            Self::Blank => Self::Text(String::new()),
            Self::Bool(b) => Self::Text(if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }),
            Self::Number(n) => Self::Text(format!("{n}")),
            Self::Text(_) => self.clone(),
            Self::Error(_) => self.clone(),
        }
    }

    /// Coerces this value to a boolean, following Excel semantics.
    ///
    /// - Blank → false
    /// - Bool → as-is
    /// - Number → 0 is false, anything else is true
    /// - Text → `#VALUE!`
    /// - Error → propagated
    pub fn to_bool(&self) -> ScalarValue {
        match self {
            Self::Blank => Self::Bool(false),
            Self::Bool(_) => self.clone(),
            Self::Number(n) => Self::Bool(*n != 0.0),
            Self::Text(s) => match s.to_uppercase().as_str() {
                "TRUE" => Self::Bool(true),
                "FALSE" => Self::Bool(false),
                _ => Self::Error(XlError::Value),
            },
            Self::Error(_) => self.clone(),
        }
    }

    /// Returns this value as a number if it is one, or `None`.
    pub fn as_number(&self) -> Option<f64> {
        match self {
            Self::Number(n) => Some(*n),
            _ => None,
        }
    }

    /// Returns this value as a bool if it is one, or `None`.
    pub fn as_bool(&self) -> Option<bool> {
        match self {
            Self::Bool(b) => Some(*b),
            _ => None,
        }
    }

    /// Returns this value as text if it is one, or `None`.
    pub fn as_text(&self) -> Option<&str> {
        match self {
            Self::Text(s) => Some(s.as_str()),
            _ => None,
        }
    }

    /// Returns the error if this is an error value.
    pub fn as_error(&self) -> Option<XlError> {
        match self {
            Self::Error(e) => Some(*e),
            _ => None,
        }
    }

    /// Returns true if this value is blank.
    pub fn is_blank(&self) -> bool {
        matches!(self, Self::Blank)
    }

    /// Returns true if this value is an error.
    pub fn is_error(&self) -> bool {
        matches!(self, Self::Error(_))
    }

    /// Compares two scalar values following Excel semantics.
    ///
    /// Excel comparison rules:
    /// - Errors propagate (first error wins)
    /// - Different types: Blank < Number < Text < Bool
    /// - Same type: natural comparison (text is case-insensitive)
    pub fn compare(&self, other: &Self) -> ScalarValue {
        // Propagate errors
        if let Self::Error(_) = self {
            return self.clone();
        }
        if let Self::Error(_) = other {
            return other.clone();
        }

        let ord = compare_values(self, other);
        Self::Number(ord as i8 as f64)
    }
}

/// Returns -1, 0, or 1 comparing two non-error scalar values following Excel rules.
///
/// Type ranking: Blank(0) < Number(1) < Text(2) < Bool(3)
pub(crate) fn compare_values(a: &ScalarValue, b: &ScalarValue) -> std::cmp::Ordering {
    fn type_rank(v: &ScalarValue) -> u8 {
        match v {
            ScalarValue::Blank => 0,
            ScalarValue::Number(_) => 1,
            ScalarValue::Text(_) => 2,
            ScalarValue::Bool(_) => 3,
            ScalarValue::Error(_) => 4, // should not reach here
        }
    }

    let rank_a = type_rank(a);
    let rank_b = type_rank(b);

    if rank_a != rank_b {
        return rank_a.cmp(&rank_b);
    }

    // Same type comparison
    match (a, b) {
        (ScalarValue::Blank, ScalarValue::Blank) => std::cmp::Ordering::Equal,
        (ScalarValue::Number(x), ScalarValue::Number(y)) => {
            x.partial_cmp(y).unwrap_or(std::cmp::Ordering::Equal)
        }
        (ScalarValue::Text(x), ScalarValue::Text(y)) => x.to_uppercase().cmp(&y.to_uppercase()),
        (ScalarValue::Bool(x), ScalarValue::Bool(y)) => x.cmp(y),
        _ => std::cmp::Ordering::Equal,
    }
}

/// The result of evaluating a formula expression.
#[derive(Debug, Clone, PartialEq)]
pub enum CalcValue {
    /// A single scalar value.
    Scalar(ScalarValue),
    // Phase 2: Array(Array2D), Reference(RangeRef)
}

impl CalcValue {
    /// Convenience constructor for a numeric value.
    pub fn number(n: f64) -> Self {
        Self::Scalar(ScalarValue::Number(n))
    }

    /// Convenience constructor for a text value.
    pub fn text(s: impl Into<String>) -> Self {
        Self::Scalar(ScalarValue::Text(s.into()))
    }

    /// Convenience constructor for a boolean value.
    pub fn bool(b: bool) -> Self {
        Self::Scalar(ScalarValue::Bool(b))
    }

    /// Convenience constructor for a blank value.
    pub fn blank() -> Self {
        Self::Scalar(ScalarValue::Blank)
    }

    /// Convenience constructor for an error value.
    pub fn error(e: XlError) -> Self {
        Self::Scalar(ScalarValue::Error(e))
    }

    /// Creates a CalcValue from a ScalarValue.
    pub fn from_scalar(s: ScalarValue) -> Self {
        Self::Scalar(s)
    }

    /// Extracts the scalar value, or returns `#VALUE!` if not scalar.
    pub fn into_scalar(self) -> ScalarValue {
        match self {
            Self::Scalar(s) => s,
        }
    }

    /// Returns a reference to the inner scalar value.
    pub fn as_scalar(&self) -> &ScalarValue {
        match self {
            Self::Scalar(s) => s,
        }
    }
}

impl From<ScalarValue> for CalcValue {
    fn from(s: ScalarValue) -> Self {
        Self::Scalar(s)
    }
}

impl From<f64> for CalcValue {
    fn from(n: f64) -> Self {
        Self::number(n)
    }
}

impl From<bool> for CalcValue {
    fn from(b: bool) -> Self {
        Self::bool(b)
    }
}

impl From<&str> for CalcValue {
    fn from(s: &str) -> Self {
        Self::text(s)
    }
}

impl From<String> for CalcValue {
    fn from(s: String) -> Self {
        Self::text(s)
    }
}

impl From<XlError> for CalcValue {
    fn from(e: XlError) -> Self {
        Self::error(e)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn blank_coercion() {
        assert_eq!(ScalarValue::Blank.to_number(), ScalarValue::Number(0.0));
        assert_eq!(
            ScalarValue::Blank.to_text(),
            ScalarValue::Text(String::new())
        );
        assert_eq!(ScalarValue::Blank.to_bool(), ScalarValue::Bool(false));
    }

    #[test]
    fn bool_coercion() {
        assert_eq!(
            ScalarValue::Bool(true).to_number(),
            ScalarValue::Number(1.0)
        );
        assert_eq!(
            ScalarValue::Bool(false).to_number(),
            ScalarValue::Number(0.0)
        );
        assert_eq!(
            ScalarValue::Bool(true).to_text(),
            ScalarValue::Text("TRUE".to_string())
        );
    }

    #[test]
    fn number_coercion() {
        assert_eq!(
            ScalarValue::Number(42.0).to_text(),
            ScalarValue::Text("42".to_string())
        );
        assert_eq!(ScalarValue::Number(0.0).to_bool(), ScalarValue::Bool(false));
        assert_eq!(ScalarValue::Number(1.0).to_bool(), ScalarValue::Bool(true));
    }

    #[test]
    fn text_to_number() {
        assert_eq!(
            ScalarValue::Text("42".to_string()).to_number(),
            ScalarValue::Number(42.0)
        );
        assert_eq!(
            ScalarValue::Text("abc".to_string()).to_number(),
            ScalarValue::Error(XlError::Value)
        );
        assert_eq!(
            ScalarValue::Text(String::new()).to_number(),
            ScalarValue::Number(0.0)
        );
    }

    #[test]
    fn error_propagation_in_coercion() {
        let err = ScalarValue::Error(XlError::Div0);
        assert_eq!(err.to_number(), ScalarValue::Error(XlError::Div0));
        assert_eq!(err.to_text(), ScalarValue::Error(XlError::Div0));
        assert_eq!(err.to_bool(), ScalarValue::Error(XlError::Div0));
    }

    #[test]
    fn xl_error_display() {
        assert_eq!(XlError::Null.to_string(), "#NULL!");
        assert_eq!(XlError::Div0.to_string(), "#DIV/0!");
        assert_eq!(XlError::Value.to_string(), "#VALUE!");
        assert_eq!(XlError::Ref.to_string(), "#REF!");
        assert_eq!(XlError::Name.to_string(), "#NAME?");
        assert_eq!(XlError::Num.to_string(), "#NUM!");
        assert_eq!(XlError::Na.to_string(), "#N/A");
    }

    #[test]
    fn xl_error_parse() {
        assert_eq!(XlError::parse("#NULL!"), Some(XlError::Null));
        assert_eq!(XlError::parse("#DIV/0!"), Some(XlError::Div0));
        assert_eq!(XlError::parse("#N/A"), Some(XlError::Na));
        assert_eq!(XlError::parse("not an error"), None);
    }

    #[test]
    fn compare_different_types() {
        // Blank < Number < Text < Bool
        assert_eq!(
            compare_values(&ScalarValue::Blank, &ScalarValue::Number(1.0)),
            std::cmp::Ordering::Less
        );
        assert_eq!(
            compare_values(
                &ScalarValue::Number(1.0),
                &ScalarValue::Text("a".to_string())
            ),
            std::cmp::Ordering::Less
        );
        assert_eq!(
            compare_values(
                &ScalarValue::Text("a".to_string()),
                &ScalarValue::Bool(false)
            ),
            std::cmp::Ordering::Less
        );
    }

    #[test]
    fn compare_same_type() {
        assert_eq!(
            compare_values(&ScalarValue::Number(1.0), &ScalarValue::Number(2.0)),
            std::cmp::Ordering::Less
        );
        assert_eq!(
            compare_values(
                &ScalarValue::Text("abc".to_string()),
                &ScalarValue::Text("ABC".to_string())
            ),
            std::cmp::Ordering::Equal
        );
    }

    #[test]
    fn calc_value_constructors() {
        assert_eq!(
            CalcValue::number(42.0).as_scalar(),
            &ScalarValue::Number(42.0)
        );
        assert_eq!(
            CalcValue::text("hello").as_scalar(),
            &ScalarValue::Text("hello".to_string())
        );
        assert_eq!(CalcValue::bool(true).as_scalar(), &ScalarValue::Bool(true));
        assert_eq!(CalcValue::blank().as_scalar(), &ScalarValue::Blank);
        assert_eq!(
            CalcValue::error(XlError::Na).as_scalar(),
            &ScalarValue::Error(XlError::Na)
        );
    }
}
