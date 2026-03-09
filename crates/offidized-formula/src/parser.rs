//! Recursive descent parser for Excel formula token streams.
//!
//! Converts a flat `Vec<Token>` (produced by the lexer) into an [`Expr`] AST.
//!
//! The grammar, from lowest to highest precedence:
//!
//! ```text
//! expression = comparison
//! comparison = concat (("=" | "<>" | "<" | "<=" | ">" | ">=") concat)*
//! concat     = addition ("&" addition)*
//! addition   = multiply (("+" | "-") multiply)*
//! multiply   = power (("*" | "/") power)*
//! power      = unary ("^" power)            // right-associative
//! unary      = ("+" | "-") unary | postfix
//! postfix    = primary "%" | primary
//! primary    = NUMBER | STRING | BOOL | ERROR
//!            | cell_ref (":" cell_ref)?
//!            | IDENT "(" args ")"
//!            | IDENT
//!            | "(" expression ")"
//! ```

use crate::ast::{BinaryOp, Expr, UnaryOp};
use crate::error::FormulaError;
use crate::reference::{column_letters_to_index, CellRef, RangeRef};
use crate::token::Token;
use crate::value::ScalarValue;

/// Parses a token stream into an expression AST.
///
/// The token stream must end with [`Token::Eof`].
pub fn parse(tokens: Vec<Token>) -> crate::error::Result<Expr> {
    let mut parser = Parser::new(tokens);
    let expr = parser.expression()?;
    if parser.peek() != &Token::Eof {
        return Err(FormulaError::Parse(format!(
            "unexpected token after expression: {:?}",
            parser.peek()
        )));
    }
    Ok(expr)
}

struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    fn new(tokens: Vec<Token>) -> Self {
        Self { tokens, pos: 0 }
    }

    /// Returns a reference to the current token without advancing.
    fn peek(&self) -> &Token {
        self.tokens.get(self.pos).unwrap_or(&Token::Eof)
    }

    /// If the current token matches `expected`, advance and return `true`.
    fn match_token(&mut self, expected: &Token) -> bool {
        if self.peek() == expected {
            self.pos += 1;
            true
        } else {
            false
        }
    }

    /// Expects the current token to match `expected`, advances, or returns an error.
    fn expect(&mut self, expected: &Token) -> crate::error::Result<()> {
        if self.peek() == expected {
            self.pos += 1;
            Ok(())
        } else {
            Err(FormulaError::Parse(format!(
                "expected {:?}, got {:?}",
                expected,
                self.peek()
            )))
        }
    }

    // ---- Grammar rules ----

    /// expression = comparison
    fn expression(&mut self) -> crate::error::Result<Expr> {
        self.comparison()
    }

    /// comparison = concat (("=" | "<>" | "<" | "<=" | ">" | ">=") concat)*
    fn comparison(&mut self) -> crate::error::Result<Expr> {
        let mut left = self.concat()?;

        loop {
            let op = match self.peek() {
                Token::Eq => BinaryOp::Eq,
                Token::Ne => BinaryOp::Ne,
                Token::Lt => BinaryOp::Lt,
                Token::Le => BinaryOp::Le,
                Token::Gt => BinaryOp::Gt,
                Token::Ge => BinaryOp::Ge,
                _ => break,
            };
            self.pos += 1;
            let right = self.concat()?;
            left = Expr::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            };
        }

        Ok(left)
    }

    /// concat = addition ("&" addition)*
    fn concat(&mut self) -> crate::error::Result<Expr> {
        let mut left = self.addition()?;

        while self.peek() == &Token::Ampersand {
            self.pos += 1;
            let right = self.addition()?;
            left = Expr::Binary {
                op: BinaryOp::Concat,
                left: Box::new(left),
                right: Box::new(right),
            };
        }

        Ok(left)
    }

    /// addition = multiply (("+" | "-") multiply)*
    fn addition(&mut self) -> crate::error::Result<Expr> {
        let mut left = self.multiply()?;

        loop {
            let op = match self.peek() {
                Token::Plus => BinaryOp::Add,
                Token::Minus => BinaryOp::Sub,
                _ => break,
            };
            self.pos += 1;
            let right = self.multiply()?;
            left = Expr::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            };
        }

        Ok(left)
    }

    /// multiply = power (("*" | "/") power)*
    fn multiply(&mut self) -> crate::error::Result<Expr> {
        let mut left = self.power()?;

        loop {
            let op = match self.peek() {
                Token::Star => BinaryOp::Mul,
                Token::Slash => BinaryOp::Div,
                _ => break,
            };
            self.pos += 1;
            let right = self.power()?;
            left = Expr::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            };
        }

        Ok(left)
    }

    /// power = unary ("^" power)  // right-associative via recursion
    fn power(&mut self) -> crate::error::Result<Expr> {
        let base = self.unary()?;

        if self.peek() == &Token::Caret {
            self.pos += 1;
            // Right-associative: recurse into power() instead of unary()
            let exp = self.power()?;
            Ok(Expr::Binary {
                op: BinaryOp::Pow,
                left: Box::new(base),
                right: Box::new(exp),
            })
        } else {
            Ok(base)
        }
    }

    /// unary = ("+" | "-") unary | postfix
    fn unary(&mut self) -> crate::error::Result<Expr> {
        match self.peek() {
            Token::Plus => {
                self.pos += 1;
                let expr = self.unary()?;
                Ok(Expr::Unary {
                    op: UnaryOp::Plus,
                    expr: Box::new(expr),
                })
            }
            Token::Minus => {
                self.pos += 1;
                let expr = self.unary()?;
                Ok(Expr::Unary {
                    op: UnaryOp::Negate,
                    expr: Box::new(expr),
                })
            }
            _ => self.postfix(),
        }
    }

    /// postfix = primary "%"*
    fn postfix(&mut self) -> crate::error::Result<Expr> {
        let mut expr = self.primary()?;

        while self.peek() == &Token::Percent {
            self.pos += 1;
            expr = Expr::Unary {
                op: UnaryOp::Percent,
                expr: Box::new(expr),
            };
        }

        Ok(expr)
    }

    /// primary = NUMBER | STRING | BOOL | ERROR
    ///         | cell_ref (":" cell_ref)?
    ///         | IDENT "(" args ")"
    ///         | IDENT
    ///         | "(" expression ")"
    fn primary(&mut self) -> crate::error::Result<Expr> {
        match self.peek().clone() {
            Token::Number(n) => {
                self.pos += 1;
                Ok(Expr::Literal(ScalarValue::Number(n)))
            }
            Token::StringLiteral(ref s) => {
                let s = s.clone();
                self.pos += 1;
                Ok(Expr::Literal(ScalarValue::Text(s)))
            }
            Token::Bool(b) => {
                self.pos += 1;
                Ok(Expr::Literal(ScalarValue::Bool(b)))
            }
            Token::Error(e) => {
                self.pos += 1;
                Ok(Expr::Literal(ScalarValue::Error(e)))
            }
            Token::CellReference {
                ref sheet,
                ref col_letters,
                row,
                col_absolute,
                row_absolute,
            } => {
                let sheet = sheet.clone();
                let col_letters = col_letters.clone();
                self.pos += 1;

                let col = column_letters_to_index(&col_letters).ok_or_else(|| {
                    FormulaError::Parse(format!("invalid column letters: {col_letters}"))
                })?;

                let cell = CellRef {
                    sheet: sheet.clone(),
                    col,
                    row,
                    col_absolute,
                    row_absolute,
                };

                // Check for range operator `:` followed by another cell ref
                if self.peek() == &Token::Colon {
                    // Peek ahead to see if next token after `:` is a cell ref
                    if let Some(Token::CellReference { .. }) = self.tokens.get(self.pos + 1) {
                        self.pos += 1; // consume `:`
                        let end = self.parse_cell_ref()?;

                        // If the start has a sheet and the end doesn't, inherit it.
                        // If neither has a sheet, that's fine too.
                        let range_sheet = sheet.or(end.sheet);

                        return Ok(Expr::RangeRef(RangeRef {
                            sheet: range_sheet,
                            start_col: col,
                            start_row: row,
                            end_col: end.col,
                            end_row: end.row,
                        }));
                    }
                }

                Ok(Expr::CellRef(cell))
            }
            Token::Ident(ref name) => {
                let name = name.clone();
                self.pos += 1;

                // Check for function call: IDENT "(" args ")"
                if self.peek() == &Token::LParen {
                    self.pos += 1; // consume `(`
                    let args = self.parse_args()?;
                    self.expect(&Token::RParen)?;
                    Ok(Expr::Function { name, args })
                } else {
                    // Named range
                    Ok(Expr::Name(name))
                }
            }
            Token::LParen => {
                self.pos += 1;
                let expr = self.expression()?;
                self.expect(&Token::RParen)?;
                Ok(expr)
            }
            _ => Err(FormulaError::Parse(format!(
                "unexpected token: {:?}",
                self.peek()
            ))),
        }
    }

    /// Parse a `Token::CellReference` at the current position into a `CellRef`.
    fn parse_cell_ref(&mut self) -> crate::error::Result<CellRef> {
        match self.peek().clone() {
            Token::CellReference {
                ref sheet,
                ref col_letters,
                row,
                col_absolute,
                row_absolute,
            } => {
                let sheet = sheet.clone();
                let col_letters = col_letters.clone();
                self.pos += 1;

                let col = column_letters_to_index(&col_letters).ok_or_else(|| {
                    FormulaError::Parse(format!("invalid column letters: {col_letters}"))
                })?;

                Ok(CellRef {
                    sheet,
                    col,
                    row,
                    col_absolute,
                    row_absolute,
                })
            }
            _ => Err(FormulaError::Parse(format!(
                "expected cell reference, got {:?}",
                self.peek()
            ))),
        }
    }

    /// Parse function argument list (comma-separated expressions).
    /// Returns an empty vec for `()`.
    fn parse_args(&mut self) -> crate::error::Result<Vec<Expr>> {
        let mut args = Vec::new();

        // Empty argument list
        if self.peek() == &Token::RParen {
            return Ok(args);
        }

        args.push(self.expression()?);

        while self.match_token(&Token::Comma) {
            args.push(self.expression()?);
        }

        Ok(args)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::lexer::tokenize;
    use crate::value::XlError;
    use pretty_assertions::assert_eq;

    /// Helper: lex + parse a formula string.
    #[allow(clippy::unwrap_used)]
    fn p(input: &str) -> Expr {
        let tokens = tokenize(input).unwrap();
        parse(tokens).unwrap()
    }

    // ---- Literals ----

    #[test]
    fn number_literal() {
        assert_eq!(p("42"), Expr::Literal(ScalarValue::Number(42.0)));
    }

    #[test]
    fn string_literal() {
        assert_eq!(
            p(r#""hello""#),
            Expr::Literal(ScalarValue::Text("hello".to_string()))
        );
    }

    #[test]
    fn bool_literal() {
        assert_eq!(p("TRUE"), Expr::Literal(ScalarValue::Bool(true)));
        assert_eq!(p("FALSE"), Expr::Literal(ScalarValue::Bool(false)));
    }

    #[test]
    fn error_literal() {
        assert_eq!(p("#N/A"), Expr::Literal(ScalarValue::Error(XlError::Na)));
    }

    // ---- Basic arithmetic ----

    #[test]
    fn addition() {
        assert_eq!(
            p("1+2"),
            Expr::Binary {
                op: BinaryOp::Add,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
            }
        );
    }

    #[test]
    fn subtraction() {
        assert_eq!(
            p("5-3"),
            Expr::Binary {
                op: BinaryOp::Sub,
                left: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
            }
        );
    }

    #[test]
    fn multiplication() {
        assert_eq!(
            p("2*3"),
            Expr::Binary {
                op: BinaryOp::Mul,
                left: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
            }
        );
    }

    #[test]
    fn division() {
        assert_eq!(
            p("6/2"),
            Expr::Binary {
                op: BinaryOp::Div,
                left: Box::new(Expr::Literal(ScalarValue::Number(6.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
            }
        );
    }

    // ---- Precedence ----

    #[test]
    fn mul_before_add() {
        // 1 + 2 * 3 => 1 + (2*3)
        assert_eq!(
            p("1+2*3"),
            Expr::Binary {
                op: BinaryOp::Add,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Binary {
                    op: BinaryOp::Mul,
                    left: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
                    right: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
                }),
            }
        );
    }

    #[test]
    fn parens_override_precedence() {
        // (1 + 2) * 3 => (1+2) * 3
        assert_eq!(
            p("(1+2)*3"),
            Expr::Binary {
                op: BinaryOp::Mul,
                left: Box::new(Expr::Binary {
                    op: BinaryOp::Add,
                    left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                    right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
                }),
                right: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
            }
        );
    }

    #[test]
    fn comparison_lower_than_add() {
        // A1 > 1 + 2 => A1 > (1+2)
        assert_eq!(
            p("1>2+3"),
            Expr::Binary {
                op: BinaryOp::Gt,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Binary {
                    op: BinaryOp::Add,
                    left: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
                    right: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
                }),
            }
        );
    }

    #[test]
    fn concat_between_comparison_and_add() {
        // "a" & "b" = "ab" => ("a" & "b") = "ab"
        // concat has higher precedence than comparison but lower than add
        assert_eq!(
            p(r#""a" & "b""#),
            Expr::Binary {
                op: BinaryOp::Concat,
                left: Box::new(Expr::Literal(ScalarValue::Text("a".to_string()))),
                right: Box::new(Expr::Literal(ScalarValue::Text("b".to_string()))),
            }
        );
    }

    // ---- Right-associative power ----

    #[test]
    fn power_right_associative() {
        // 2^3^4 => 2^(3^4), NOT (2^3)^4
        assert_eq!(
            p("2^3^4"),
            Expr::Binary {
                op: BinaryOp::Pow,
                left: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
                right: Box::new(Expr::Binary {
                    op: BinaryOp::Pow,
                    left: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
                    right: Box::new(Expr::Literal(ScalarValue::Number(4.0))),
                }),
            }
        );
    }

    #[test]
    fn power_higher_than_mul() {
        // 2*3^4 => 2*(3^4)
        assert_eq!(
            p("2*3^4"),
            Expr::Binary {
                op: BinaryOp::Mul,
                left: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
                right: Box::new(Expr::Binary {
                    op: BinaryOp::Pow,
                    left: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
                    right: Box::new(Expr::Literal(ScalarValue::Number(4.0))),
                }),
            }
        );
    }

    // ---- Unary operators ----

    #[test]
    fn unary_negate() {
        assert_eq!(
            p("-5"),
            Expr::Unary {
                op: UnaryOp::Negate,
                expr: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
            }
        );
    }

    #[test]
    fn unary_plus() {
        assert_eq!(
            p("+5"),
            Expr::Unary {
                op: UnaryOp::Plus,
                expr: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
            }
        );
    }

    #[test]
    fn double_negate() {
        assert_eq!(
            p("--5"),
            Expr::Unary {
                op: UnaryOp::Negate,
                expr: Box::new(Expr::Unary {
                    op: UnaryOp::Negate,
                    expr: Box::new(Expr::Literal(ScalarValue::Number(5.0))),
                }),
            }
        );
    }

    // ---- Percent ----

    #[test]
    fn percent_postfix() {
        assert_eq!(
            p("50%"),
            Expr::Unary {
                op: UnaryOp::Percent,
                expr: Box::new(Expr::Literal(ScalarValue::Number(50.0))),
            }
        );
    }

    #[test]
    fn percent_in_expression() {
        // 50% + 1 => (50%) + 1
        assert_eq!(
            p("50%+1"),
            Expr::Binary {
                op: BinaryOp::Add,
                left: Box::new(Expr::Unary {
                    op: UnaryOp::Percent,
                    expr: Box::new(Expr::Literal(ScalarValue::Number(50.0))),
                }),
                right: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
            }
        );
    }

    // ---- Cell references ----

    #[test]
    fn simple_cell_ref() {
        assert_eq!(
            p("A1"),
            Expr::CellRef(CellRef {
                sheet: None,
                col: 1,
                row: 1,
                col_absolute: false,
                row_absolute: false,
            })
        );
    }

    #[test]
    fn absolute_cell_ref() {
        assert_eq!(
            p("$B$3"),
            Expr::CellRef(CellRef {
                sheet: None,
                col: 2,
                row: 3,
                col_absolute: true,
                row_absolute: true,
            })
        );
    }

    #[test]
    fn cell_ref_with_sheet() {
        assert_eq!(
            p("Sheet1!A1"),
            Expr::CellRef(CellRef {
                sheet: Some("Sheet1".to_string()),
                col: 1,
                row: 1,
                col_absolute: false,
                row_absolute: false,
            })
        );
    }

    #[test]
    fn cell_ref_with_quoted_sheet() {
        assert_eq!(
            p("'My Sheet'!C5"),
            Expr::CellRef(CellRef {
                sheet: Some("My Sheet".to_string()),
                col: 3,
                row: 5,
                col_absolute: false,
                row_absolute: false,
            })
        );
    }

    // ---- Ranges ----

    #[test]
    fn simple_range() {
        assert_eq!(
            p("A1:C5"),
            Expr::RangeRef(RangeRef {
                sheet: None,
                start_col: 1,
                start_row: 1,
                end_col: 3,
                end_row: 5,
            })
        );
    }

    #[test]
    fn range_with_sheet() {
        assert_eq!(
            p("Sheet1!A1:C5"),
            Expr::RangeRef(RangeRef {
                sheet: Some("Sheet1".to_string()),
                start_col: 1,
                start_row: 1,
                end_col: 3,
                end_row: 5,
            })
        );
    }

    #[test]
    fn range_with_absolute_refs() {
        assert_eq!(
            p("$A$1:$C$5"),
            Expr::RangeRef(RangeRef {
                sheet: None,
                start_col: 1,
                start_row: 1,
                end_col: 3,
                end_row: 5,
            })
        );
    }

    // ---- Function calls ----

    #[test]
    fn function_no_args() {
        assert_eq!(
            p("NOW()"),
            Expr::Function {
                name: "NOW".to_string(),
                args: vec![],
            }
        );
    }

    #[test]
    fn function_one_arg() {
        assert_eq!(
            p("ABS(-1)"),
            Expr::Function {
                name: "ABS".to_string(),
                args: vec![Expr::Unary {
                    op: UnaryOp::Negate,
                    expr: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                }],
            }
        );
    }

    #[test]
    fn function_multiple_args() {
        assert_eq!(
            p("IF(TRUE,1,2)"),
            Expr::Function {
                name: "IF".to_string(),
                args: vec![
                    Expr::Literal(ScalarValue::Bool(true)),
                    Expr::Literal(ScalarValue::Number(1.0)),
                    Expr::Literal(ScalarValue::Number(2.0)),
                ],
            }
        );
    }

    #[test]
    fn function_with_range_arg() {
        assert_eq!(
            p("SUM(A1:B2)"),
            Expr::Function {
                name: "SUM".to_string(),
                args: vec![Expr::RangeRef(RangeRef {
                    sheet: None,
                    start_col: 1,
                    start_row: 1,
                    end_col: 2,
                    end_row: 2,
                })],
            }
        );
    }

    #[test]
    fn nested_functions() {
        assert_eq!(
            p("SUM(ABS(1),ABS(2))"),
            Expr::Function {
                name: "SUM".to_string(),
                args: vec![
                    Expr::Function {
                        name: "ABS".to_string(),
                        args: vec![Expr::Literal(ScalarValue::Number(1.0))],
                    },
                    Expr::Function {
                        name: "ABS".to_string(),
                        args: vec![Expr::Literal(ScalarValue::Number(2.0))],
                    },
                ],
            }
        );
    }

    // ---- Named ranges ----

    #[test]
    fn named_range() {
        assert_eq!(p("MyRange"), Expr::Name("MyRange".to_string()));
    }

    // ---- Nested expressions ----

    #[test]
    fn deeply_nested_parens() {
        assert_eq!(
            p("((1+2))"),
            Expr::Binary {
                op: BinaryOp::Add,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
            }
        );
    }

    // ---- Complex formulas ----

    #[test]
    fn complex_formula() {
        // =IF(A1>0,SUM(B1:B10),0)
        let expr = p("IF(A1>0,SUM(B1:B10),0)");
        assert!(matches!(
            &expr,
            Expr::Function { name, args }
            if name == "IF"
                && args.len() == 3
                && matches!(&args[0], Expr::Binary { op: BinaryOp::Gt, .. })
                && matches!(&args[1], Expr::Function { name: inner_name, args: inner_args }
                    if inner_name == "SUM" && inner_args.len() == 1)
                && args[2] == Expr::Literal(ScalarValue::Number(0.0))
        ));
    }

    #[test]
    fn comparison_operators() {
        assert_eq!(
            p("1=1"),
            Expr::Binary {
                op: BinaryOp::Eq,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
            }
        );
        assert_eq!(
            p("1<>2"),
            Expr::Binary {
                op: BinaryOp::Ne,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
            }
        );
        assert_eq!(
            p("1<2"),
            Expr::Binary {
                op: BinaryOp::Lt,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
            }
        );
        assert_eq!(
            p("1<=2"),
            Expr::Binary {
                op: BinaryOp::Le,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
            }
        );
        assert_eq!(
            p("1>=2"),
            Expr::Binary {
                op: BinaryOp::Ge,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
            }
        );
    }

    // ---- Associativity ----

    #[test]
    fn left_associative_addition() {
        // 1 - 2 - 3 => (1 - 2) - 3
        assert_eq!(
            p("1-2-3"),
            Expr::Binary {
                op: BinaryOp::Sub,
                left: Box::new(Expr::Binary {
                    op: BinaryOp::Sub,
                    left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                    right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
                }),
                right: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
            }
        );
    }

    #[test]
    fn left_associative_multiplication() {
        // 2 * 3 * 4 => (2 * 3) * 4
        assert_eq!(
            p("2*3*4"),
            Expr::Binary {
                op: BinaryOp::Mul,
                left: Box::new(Expr::Binary {
                    op: BinaryOp::Mul,
                    left: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
                    right: Box::new(Expr::Literal(ScalarValue::Number(3.0))),
                }),
                right: Box::new(Expr::Literal(ScalarValue::Number(4.0))),
            }
        );
    }

    // ---- Error cases ----

    #[test]
    #[allow(clippy::unwrap_used)]
    fn parse_error_unexpected_token() {
        let tokens = tokenize(")").unwrap();
        assert!(parse(tokens).is_err());
    }

    #[test]
    #[allow(clippy::unwrap_used)]
    fn parse_error_trailing_tokens() {
        let tokens = tokenize("1 2").unwrap();
        assert!(parse(tokens).is_err());
    }

    #[test]
    #[allow(clippy::unwrap_used)]
    fn parse_error_unclosed_paren() {
        let tokens = tokenize("(1+2").unwrap();
        assert!(parse(tokens).is_err());
    }

    #[test]
    fn unary_negate_in_expression() {
        // 1 + -2 => 1 + (-(2))
        assert_eq!(
            p("1+-2"),
            Expr::Binary {
                op: BinaryOp::Add,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Unary {
                    op: UnaryOp::Negate,
                    expr: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
                }),
            }
        );
    }

    #[test]
    fn string_concat() {
        assert_eq!(
            p(r#""a" & "b" & "c""#),
            Expr::Binary {
                op: BinaryOp::Concat,
                left: Box::new(Expr::Binary {
                    op: BinaryOp::Concat,
                    left: Box::new(Expr::Literal(ScalarValue::Text("a".to_string()))),
                    right: Box::new(Expr::Literal(ScalarValue::Text("b".to_string()))),
                }),
                right: Box::new(Expr::Literal(ScalarValue::Text("c".to_string()))),
            }
        );
    }

    #[test]
    fn cell_ref_in_arithmetic() {
        assert_eq!(
            p("A1+B2"),
            Expr::Binary {
                op: BinaryOp::Add,
                left: Box::new(Expr::CellRef(CellRef {
                    sheet: None,
                    col: 1,
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                })),
                right: Box::new(Expr::CellRef(CellRef {
                    sheet: None,
                    col: 2,
                    row: 2,
                    col_absolute: false,
                    row_absolute: false,
                })),
            }
        );
    }

    #[test]
    fn leading_equals() {
        // `=1+2` should parse the same as `1+2`
        assert_eq!(
            p("=1+2"),
            Expr::Binary {
                op: BinaryOp::Add,
                left: Box::new(Expr::Literal(ScalarValue::Number(1.0))),
                right: Box::new(Expr::Literal(ScalarValue::Number(2.0))),
            }
        );
    }
}
