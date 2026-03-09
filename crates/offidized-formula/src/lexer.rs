//! Hand-written scanner for Excel formula strings.
//!
//! Converts a formula string like `=SUM(A1:B2, 3.14)` into a flat `Vec<Token>`.

use crate::error::FormulaError;
use crate::token::Token;
use crate::value::XlError;

/// Tokenizes an Excel formula string into a sequence of tokens.
///
/// The leading `=` is optional and will be skipped if present.
/// The returned token stream always ends with [`Token::Eof`].
pub fn tokenize(input: &str) -> crate::error::Result<Vec<Token>> {
    let mut lexer = Lexer::new(input);
    lexer.scan_all()
}

struct Lexer<'a> {
    /// The full input as bytes.
    chars: &'a [u8],
    /// Current read position.
    pos: usize,
}

impl<'a> Lexer<'a> {
    fn new(input: &'a str) -> Self {
        Self {
            chars: input.as_bytes(),
            pos: 0,
        }
    }

    /// Returns the current byte without advancing, or `None` at end.
    fn peek(&self) -> Option<u8> {
        self.chars.get(self.pos).copied()
    }

    /// Returns the byte at `pos + offset`, or `None`.
    fn peek_at(&self, offset: usize) -> Option<u8> {
        self.chars.get(self.pos + offset).copied()
    }

    /// Advances and returns the current byte.
    fn advance(&mut self) -> Option<u8> {
        let ch = self.chars.get(self.pos).copied();
        if ch.is_some() {
            self.pos += 1;
        }
        ch
    }

    /// Returns `true` if there are no more bytes.
    fn at_end(&self) -> bool {
        self.pos >= self.chars.len()
    }

    fn skip_whitespace(&mut self) {
        while let Some(ch) = self.peek() {
            if ch == b' ' || ch == b'\t' || ch == b'\r' || ch == b'\n' {
                self.pos += 1;
            } else {
                break;
            }
        }
    }

    /// Peek past whitespace to see the next non-whitespace byte.
    fn peek_past_whitespace(&self) -> Option<u8> {
        let mut i = self.pos;
        while i < self.chars.len() {
            let ch = self.chars[i];
            if ch == b' ' || ch == b'\t' || ch == b'\r' || ch == b'\n' {
                i += 1;
            } else {
                return Some(ch);
            }
        }
        None
    }

    fn scan_all(&mut self) -> crate::error::Result<Vec<Token>> {
        let mut tokens = Vec::new();

        // Skip leading `=` if present.
        self.skip_whitespace();
        if self.peek() == Some(b'=') {
            self.pos += 1;
        }

        loop {
            self.skip_whitespace();

            if self.at_end() {
                tokens.push(Token::Eof);
                return Ok(tokens);
            }

            let ch = match self.peek() {
                Some(c) => c,
                None => {
                    tokens.push(Token::Eof);
                    return Ok(tokens);
                }
            };

            match ch {
                b'+' => {
                    self.pos += 1;
                    tokens.push(Token::Plus);
                }
                b'-' => {
                    self.pos += 1;
                    tokens.push(Token::Minus);
                }
                b'*' => {
                    self.pos += 1;
                    tokens.push(Token::Star);
                }
                b'/' => {
                    self.pos += 1;
                    tokens.push(Token::Slash);
                }
                b'^' => {
                    self.pos += 1;
                    tokens.push(Token::Caret);
                }
                b'%' => {
                    self.pos += 1;
                    tokens.push(Token::Percent);
                }
                b'&' => {
                    self.pos += 1;
                    tokens.push(Token::Ampersand);
                }
                b':' => {
                    self.pos += 1;
                    tokens.push(Token::Colon);
                }
                b',' => {
                    self.pos += 1;
                    tokens.push(Token::Comma);
                }
                b'(' => {
                    self.pos += 1;
                    tokens.push(Token::LParen);
                }
                b')' => {
                    self.pos += 1;
                    tokens.push(Token::RParen);
                }
                b'=' => {
                    self.pos += 1;
                    tokens.push(Token::Eq);
                }
                b'<' => {
                    self.pos += 1;
                    if self.peek() == Some(b'>') {
                        self.pos += 1;
                        tokens.push(Token::Ne);
                    } else if self.peek() == Some(b'=') {
                        self.pos += 1;
                        tokens.push(Token::Le);
                    } else {
                        tokens.push(Token::Lt);
                    }
                }
                b'>' => {
                    self.pos += 1;
                    if self.peek() == Some(b'=') {
                        self.pos += 1;
                        tokens.push(Token::Ge);
                    } else {
                        tokens.push(Token::Gt);
                    }
                }
                b'"' => {
                    let tok = self.scan_string()?;
                    tokens.push(tok);
                }
                b'#' => {
                    let tok = self.scan_error_literal()?;
                    tokens.push(tok);
                }
                b'$' => {
                    // Must be start of a cell reference like $A$1
                    let tok = self.scan_dollar_cell_ref(None)?;
                    tokens.push(tok);
                }
                b'\'' => {
                    // Quoted sheet name prefix: 'My Sheet'!A1
                    let sheet = self.scan_quoted_sheet()?;
                    self.skip_whitespace();
                    let tok = self.scan_after_sheet(sheet)?;
                    tokens.push(tok);
                }
                b'.' | b'0'..=b'9' => {
                    let tok = self.scan_number()?;
                    tokens.push(tok);
                }
                b'A'..=b'Z' | b'a'..=b'z' | b'_' => {
                    let tok = self.scan_word()?;
                    tokens.push(tok);
                }
                other => {
                    return Err(FormulaError::Lex(format!(
                        "unexpected character: '{}'",
                        other as char
                    )));
                }
            }
        }
    }

    /// Scans a string literal: `"hello ""world"""` -> `hello "world"`.
    fn scan_string(&mut self) -> crate::error::Result<Token> {
        // consume opening quote
        self.pos += 1;
        let mut value = String::new();
        loop {
            match self.advance() {
                None => {
                    return Err(FormulaError::Lex("unterminated string literal".to_string()));
                }
                Some(b'"') => {
                    // Check for escaped quote (`""`)
                    if self.peek() == Some(b'"') {
                        self.pos += 1;
                        value.push('"');
                    } else {
                        return Ok(Token::StringLiteral(value));
                    }
                }
                Some(ch) => {
                    value.push(ch as char);
                }
            }
        }
    }

    /// Scans an error literal starting with `#`.
    fn scan_error_literal(&mut self) -> crate::error::Result<Token> {
        let start = self.pos;
        // Consume `#`
        self.pos += 1;

        // Consume until we hit `!` or `?` (the error terminator) or run out of chars.
        loop {
            match self.peek() {
                Some(b'!' | b'?') => {
                    self.pos += 1;
                    break;
                }
                Some(b'A'..=b'Z' | b'a'..=b'z' | b'/' | b'0'..=b'9') => {
                    self.pos += 1;
                }
                _ => break,
            }
        }

        let text = std::str::from_utf8(&self.chars[start..self.pos])
            .map_err(|e| FormulaError::Lex(format!("invalid UTF-8 in error literal: {e}")))?;

        // Uppercase for matching
        let upper = text.to_uppercase();
        match XlError::parse(&upper) {
            Some(err) => Ok(Token::Error(err)),
            None => Err(FormulaError::Lex(format!("unknown error literal: {text}"))),
        }
    }

    /// Scans a quoted sheet name like `'My Sheet'` and expects `!` after.
    fn scan_quoted_sheet(&mut self) -> crate::error::Result<String> {
        // consume opening quote
        self.pos += 1;
        let mut name = String::new();
        loop {
            match self.advance() {
                None => {
                    return Err(FormulaError::Lex(
                        "unterminated quoted sheet name".to_string(),
                    ));
                }
                Some(b'\'') => {
                    // Check for escaped single quote (`''`)
                    if self.peek() == Some(b'\'') {
                        self.pos += 1;
                        name.push('\'');
                    } else {
                        break;
                    }
                }
                Some(ch) => {
                    name.push(ch as char);
                }
            }
        }
        // Expect `!` after the closing quote
        if self.peek() != Some(b'!') {
            return Err(FormulaError::Lex(
                "expected '!' after quoted sheet name".to_string(),
            ));
        }
        self.pos += 1; // consume `!`
        Ok(name)
    }

    /// Scans a number: integer, decimal, or scientific notation.
    ///
    /// A number must start with a digit or `.` followed by a digit.
    fn scan_number(&mut self) -> crate::error::Result<Token> {
        let start = self.pos;

        // Consume leading digits
        while let Some(b'0'..=b'9') = self.peek() {
            self.pos += 1;
        }

        // Decimal part
        if self.peek() == Some(b'.') {
            self.pos += 1;
            while let Some(b'0'..=b'9') = self.peek() {
                self.pos += 1;
            }
        }

        // Scientific notation: E or e, optionally followed by + or -
        if let Some(b'E' | b'e') = self.peek() {
            // Only consume if followed by a digit or sign+digit
            let next = self.peek_at(1);
            let is_sci = match next {
                Some(b'0'..=b'9') => true,
                Some(b'+' | b'-') => matches!(self.peek_at(2), Some(b'0'..=b'9')),
                _ => false,
            };
            if is_sci {
                self.pos += 1; // consume E/e
                if let Some(b'+' | b'-') = self.peek() {
                    self.pos += 1;
                }
                while let Some(b'0'..=b'9') = self.peek() {
                    self.pos += 1;
                }
            }
        }

        let text = std::str::from_utf8(&self.chars[start..self.pos])
            .map_err(|e| FormulaError::Lex(format!("invalid UTF-8 in number: {e}")))?;

        let value: f64 = text
            .parse()
            .map_err(|e| FormulaError::Lex(format!("invalid number '{text}': {e}")))?;

        Ok(Token::Number(value))
    }

    /// Scans a word that starts with a letter or underscore.
    ///
    /// This could be:
    /// - A boolean literal (`TRUE`, `FALSE`) -- unless followed by `(`
    /// - A cell reference like `A1`, `AB123`
    /// - A cell reference with absolute row like `A$1`
    /// - A sheet-qualified cell reference like `Sheet1!A1`
    /// - A function name like `SUM`
    /// - A named range identifier
    fn scan_word(&mut self) -> crate::error::Result<Token> {
        // First, consume letters only (for potential cell reference column part).
        let letter_start = self.pos;
        while let Some(ch) = self.peek() {
            if ch.is_ascii_alphabetic() {
                self.pos += 1;
            } else {
                break;
            }
        }
        let letter_end = self.pos;
        let has_letters = letter_end > letter_start;

        // Check if after letters we have `$` + digit (absolute row cell ref like `A$1`)
        // or digits (normal cell ref like `A1`), but only if all consumed so far is letters.
        if has_letters {
            let after = self.peek();

            // Check: letters followed by `$` + digit => cell ref with absolute row
            if after == Some(b'$') {
                if let Some(b'0'..=b'9') = self.peek_at(1) {
                    let col_letters_bytes = &self.chars[letter_start..letter_end];
                    let col_letters = std::str::from_utf8(col_letters_bytes)
                        .map_err(|e| FormulaError::Lex(format!("invalid UTF-8: {e}")))?
                        .to_uppercase();

                    // Check for `!` prefix first -- letters then `!` => sheet prefix
                    // (not applicable here since we already checked for `$` after letters)

                    // It's a cell ref with absolute row
                    self.pos += 1; // consume `$`
                    let row = self.scan_row_digits()?;

                    // But wait -- we need to check if `letters!` case happened already.
                    // At this point the path is: letters, `$`, digits. That's `A$1` style.
                    return Ok(Token::CellReference {
                        sheet: None,
                        col_letters,
                        row,
                        col_absolute: false,
                        row_absolute: true,
                    });
                }
            }

            // Check: letters followed by digits => possible cell ref
            if let Some(b'0'..=b'9') = after {
                // Consume the rest of the identifier (letters+digits+underscore)
                // to get the full word, then decide.
                let digit_start = self.pos;
                while let Some(b'0'..=b'9') = self.peek() {
                    self.pos += 1;
                }
                let digit_end = self.pos;

                // If followed by more letters/underscore, it is an identifier, not a cell ref.
                if let Some(ch) = self.peek() {
                    if ch.is_ascii_alphabetic() || ch == b'_' {
                        // Continue consuming as identifier
                        while let Some(ch) = self.peek() {
                            if ch.is_ascii_alphanumeric() || ch == b'_' || ch == b'.' {
                                self.pos += 1;
                            } else {
                                break;
                            }
                        }
                        let text = std::str::from_utf8(&self.chars[letter_start..self.pos])
                            .map_err(|e| {
                                FormulaError::Lex(format!("invalid UTF-8 in identifier: {e}"))
                            })?;
                        return self.classify_identifier(text);
                    }
                }

                // Check if followed by `!` => sheet prefix (e.g. `Sheet1!A1`)
                if self.peek() == Some(b'!') {
                    let text =
                        std::str::from_utf8(&self.chars[letter_start..digit_end]).map_err(|e| {
                            FormulaError::Lex(format!("invalid UTF-8 in identifier: {e}"))
                        })?;
                    let sheet = text.to_string();
                    self.pos += 1; // consume `!`
                    return self.scan_after_sheet(sheet);
                }

                // Check if followed by `(` => function name.
                if self.peek_past_whitespace() == Some(b'(') {
                    let text =
                        std::str::from_utf8(&self.chars[letter_start..digit_end]).map_err(|e| {
                            FormulaError::Lex(format!("invalid UTF-8 in identifier: {e}"))
                        })?;
                    return Ok(Token::Ident(text.to_uppercase()));
                }

                // It is a cell reference.
                let col_letters = std::str::from_utf8(&self.chars[letter_start..letter_end])
                    .map_err(|e| FormulaError::Lex(format!("invalid UTF-8: {e}")))?
                    .to_uppercase();
                let row_str = std::str::from_utf8(&self.chars[digit_start..digit_end])
                    .map_err(|e| FormulaError::Lex(format!("invalid UTF-8: {e}")))?;
                let row: u32 = row_str
                    .parse()
                    .map_err(|e| FormulaError::Lex(format!("invalid row: {e}")))?;
                if row == 0 {
                    return Err(FormulaError::Lex("row number cannot be 0".to_string()));
                }
                return Ok(Token::CellReference {
                    sheet: None,
                    col_letters,
                    row,
                    col_absolute: false,
                    row_absolute: false,
                });
            }
        }

        // Continue consuming the rest as a general identifier (letters, digits, underscore, dot).
        while let Some(ch) = self.peek() {
            if ch.is_ascii_alphanumeric() || ch == b'_' || ch == b'.' {
                self.pos += 1;
            } else {
                break;
            }
        }

        let text = std::str::from_utf8(&self.chars[letter_start..self.pos])
            .map_err(|e| FormulaError::Lex(format!("invalid UTF-8 in identifier: {e}")))?;

        // Check for sheet prefix: identifier followed by `!`
        if self.peek() == Some(b'!') {
            let sheet = text.to_string();
            self.pos += 1; // consume `!`
            return self.scan_after_sheet(sheet);
        }

        self.classify_identifier(text)
    }

    /// Classify a fully-consumed identifier as boolean, function name, or named range.
    fn classify_identifier(&mut self, text: &str) -> crate::error::Result<Token> {
        let upper = text.to_uppercase();

        // Check if followed by `(` => function/ident
        if self.peek_past_whitespace() == Some(b'(') {
            return Ok(Token::Ident(upper));
        }

        // Check for boolean literals (only when NOT followed by `(`)
        if upper == "TRUE" {
            return Ok(Token::Bool(true));
        }
        if upper == "FALSE" {
            return Ok(Token::Bool(false));
        }

        // Named range or other identifier
        Ok(Token::Ident(text.to_string()))
    }

    /// After a sheet prefix (e.g. `Sheet1!` or `'My Sheet'!`), scan the cell reference.
    /// This handles `$A$1`, `A1`, `$A1`, `A$1` forms.
    fn scan_after_sheet(&mut self, sheet: String) -> crate::error::Result<Token> {
        // May start with `$` for absolute column
        let col_absolute = if self.peek() == Some(b'$') {
            self.pos += 1;
            true
        } else {
            false
        };

        // Consume column letters
        let letter_start = self.pos;
        while let Some(ch) = self.peek() {
            if ch.is_ascii_alphabetic() {
                self.pos += 1;
            } else {
                break;
            }
        }
        let letter_end = self.pos;

        if letter_start == letter_end {
            return Err(FormulaError::Lex(
                "expected column letters after sheet prefix".to_string(),
            ));
        }

        let col_letters = std::str::from_utf8(&self.chars[letter_start..letter_end])
            .map_err(|e| FormulaError::Lex(format!("invalid UTF-8: {e}")))?
            .to_uppercase();

        // Possibly `$` for absolute row
        let row_absolute = if self.peek() == Some(b'$') {
            self.pos += 1;
            true
        } else {
            false
        };

        let row = self.scan_row_digits()?;

        Ok(Token::CellReference {
            sheet: Some(sheet),
            col_letters,
            row,
            col_absolute,
            row_absolute,
        })
    }

    /// Scans a cell reference that starts with `$` (absolute column).
    fn scan_dollar_cell_ref(&mut self, sheet: Option<String>) -> crate::error::Result<Token> {
        // Consume `$`
        self.pos += 1;
        let col_absolute = true;

        // Consume column letters
        let letter_start = self.pos;
        while let Some(ch) = self.peek() {
            if ch.is_ascii_alphabetic() {
                self.pos += 1;
            } else {
                break;
            }
        }
        let letter_end = self.pos;

        if letter_start == letter_end {
            return Err(FormulaError::Lex(
                "expected column letters in cell reference".to_string(),
            ));
        }

        let col_letters = std::str::from_utf8(&self.chars[letter_start..letter_end])
            .map_err(|e| FormulaError::Lex(format!("invalid UTF-8: {e}")))?
            .to_uppercase();

        // Possibly `$` for absolute row
        let row_absolute = if self.peek() == Some(b'$') {
            self.pos += 1;
            true
        } else {
            false
        };

        let row = self.scan_row_digits()?;

        Ok(Token::CellReference {
            sheet,
            col_letters,
            row,
            col_absolute,
            row_absolute,
        })
    }

    /// Consumes and returns a row number (one or more digits).
    fn scan_row_digits(&mut self) -> crate::error::Result<u32> {
        let digit_start = self.pos;
        while let Some(b'0'..=b'9') = self.peek() {
            self.pos += 1;
        }
        let digit_end = self.pos;

        if digit_start == digit_end {
            return Err(FormulaError::Lex(
                "expected row number in cell reference".to_string(),
            ));
        }

        let row_str = std::str::from_utf8(&self.chars[digit_start..digit_end])
            .map_err(|e| FormulaError::Lex(format!("invalid UTF-8: {e}")))?;
        let row: u32 = row_str
            .parse()
            .map_err(|e| FormulaError::Lex(format!("invalid row number '{row_str}': {e}")))?;

        if row == 0 {
            return Err(FormulaError::Lex("row number cannot be 0".to_string()));
        }

        Ok(row)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    /// Helper: tokenize and strip the trailing Eof for easier assertion.
    #[allow(clippy::unwrap_used)]
    fn toks(input: &str) -> Vec<Token> {
        let mut v = tokenize(input).unwrap();
        // Remove trailing Eof for test convenience
        if v.last() == Some(&Token::Eof) {
            v.pop();
        }
        v
    }

    // ---- Numbers ----

    #[test]
    fn integer() {
        assert_eq!(toks("42"), vec![Token::Number(42.0)]);
    }

    #[test]
    fn decimal() {
        assert_eq!(toks("3.15"), vec![Token::Number(3.15)]);
    }

    #[test]
    fn leading_dot() {
        assert_eq!(toks(".5"), vec![Token::Number(0.5)]);
    }

    #[test]
    fn scientific_notation() {
        assert_eq!(toks("1.5E3"), vec![Token::Number(1500.0)]);
    }

    #[test]
    fn scientific_notation_negative_exponent() {
        assert_eq!(toks("1.5E-3"), vec![Token::Number(0.0015)]);
    }

    #[test]
    fn scientific_notation_positive_exponent() {
        assert_eq!(toks("2e+4"), vec![Token::Number(20000.0)]);
    }

    #[test]
    fn zero() {
        assert_eq!(toks("0"), vec![Token::Number(0.0)]);
    }

    // ---- Strings ----

    #[test]
    fn simple_string() {
        assert_eq!(
            toks(r#""hello""#),
            vec![Token::StringLiteral("hello".to_string())]
        );
    }

    #[test]
    fn empty_string() {
        assert_eq!(toks(r#""""#), vec![Token::StringLiteral(String::new())]);
    }

    #[test]
    fn string_with_escaped_quote() {
        assert_eq!(
            toks(r#""say ""hi""  ""#),
            vec![Token::StringLiteral("say \"hi\"  ".to_string())]
        );
    }

    #[test]
    fn unterminated_string() {
        assert!(tokenize(r#""hello"#).is_err());
    }

    // ---- Booleans ----

    #[test]
    fn boolean_true() {
        assert_eq!(toks("TRUE"), vec![Token::Bool(true)]);
    }

    #[test]
    fn boolean_false() {
        assert_eq!(toks("FALSE"), vec![Token::Bool(false)]);
    }

    #[test]
    fn boolean_case_insensitive() {
        assert_eq!(toks("true"), vec![Token::Bool(true)]);
        assert_eq!(toks("False"), vec![Token::Bool(false)]);
    }

    #[test]
    fn boolean_followed_by_paren_is_function() {
        // TRUE() should be Ident("TRUE"), not Bool(true)
        assert_eq!(
            toks("TRUE()"),
            vec![
                Token::Ident("TRUE".to_string()),
                Token::LParen,
                Token::RParen
            ]
        );
    }

    // ---- Error literals ----

    #[test]
    fn error_null() {
        assert_eq!(toks("#NULL!"), vec![Token::Error(XlError::Null)]);
    }

    #[test]
    fn error_div0() {
        assert_eq!(toks("#DIV/0!"), vec![Token::Error(XlError::Div0)]);
    }

    #[test]
    fn error_value() {
        assert_eq!(toks("#VALUE!"), vec![Token::Error(XlError::Value)]);
    }

    #[test]
    fn error_ref() {
        assert_eq!(toks("#REF!"), vec![Token::Error(XlError::Ref)]);
    }

    #[test]
    fn error_name() {
        assert_eq!(toks("#NAME?"), vec![Token::Error(XlError::Name)]);
    }

    #[test]
    fn error_num() {
        assert_eq!(toks("#NUM!"), vec![Token::Error(XlError::Num)]);
    }

    #[test]
    fn error_na() {
        assert_eq!(toks("#N/A"), vec![Token::Error(XlError::Na)]);
    }

    // ---- Cell references ----

    #[test]
    fn simple_cell_ref() {
        assert_eq!(
            toks("A1"),
            vec![Token::CellReference {
                sheet: None,
                col_letters: "A".to_string(),
                row: 1,
                col_absolute: false,
                row_absolute: false,
            }]
        );
    }

    #[test]
    fn cell_ref_multi_letter() {
        assert_eq!(
            toks("AB123"),
            vec![Token::CellReference {
                sheet: None,
                col_letters: "AB".to_string(),
                row: 123,
                col_absolute: false,
                row_absolute: false,
            }]
        );
    }

    #[test]
    fn cell_ref_absolute_col() {
        assert_eq!(
            toks("$A1"),
            vec![Token::CellReference {
                sheet: None,
                col_letters: "A".to_string(),
                row: 1,
                col_absolute: true,
                row_absolute: false,
            }]
        );
    }

    #[test]
    fn cell_ref_absolute_row() {
        assert_eq!(
            toks("A$1"),
            vec![Token::CellReference {
                sheet: None,
                col_letters: "A".to_string(),
                row: 1,
                col_absolute: false,
                row_absolute: true,
            }]
        );
    }

    #[test]
    fn cell_ref_fully_absolute() {
        assert_eq!(
            toks("$A$1"),
            vec![Token::CellReference {
                sheet: None,
                col_letters: "A".to_string(),
                row: 1,
                col_absolute: true,
                row_absolute: true,
            }]
        );
    }

    #[test]
    fn cell_ref_e3_is_cell_not_number() {
        // E3 should be a cell ref, not scientific notation
        assert_eq!(
            toks("E3"),
            vec![Token::CellReference {
                sheet: None,
                col_letters: "E".to_string(),
                row: 3,
                col_absolute: false,
                row_absolute: false,
            }]
        );
    }

    // ---- Sheet prefixes ----

    #[test]
    fn sheet_prefix_simple() {
        assert_eq!(
            toks("Sheet1!A1"),
            vec![Token::CellReference {
                sheet: Some("Sheet1".to_string()),
                col_letters: "A".to_string(),
                row: 1,
                col_absolute: false,
                row_absolute: false,
            }]
        );
    }

    #[test]
    fn sheet_prefix_quoted() {
        assert_eq!(
            toks("'My Sheet'!A1"),
            vec![Token::CellReference {
                sheet: Some("My Sheet".to_string()),
                col_letters: "A".to_string(),
                row: 1,
                col_absolute: false,
                row_absolute: false,
            }]
        );
    }

    #[test]
    fn sheet_prefix_with_absolute_cell() {
        assert_eq!(
            toks("Sheet2!$B$3"),
            vec![Token::CellReference {
                sheet: Some("Sheet2".to_string()),
                col_letters: "B".to_string(),
                row: 3,
                col_absolute: true,
                row_absolute: true,
            }]
        );
    }

    // ---- Function names vs cell refs ----

    #[test]
    fn function_name_sum() {
        assert_eq!(
            toks("SUM(A1)"),
            vec![
                Token::Ident("SUM".to_string()),
                Token::LParen,
                Token::CellReference {
                    sheet: None,
                    col_letters: "A".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::RParen,
            ]
        );
    }

    #[test]
    fn function_name_if() {
        assert_eq!(
            toks("IF(A1,1,2)"),
            vec![
                Token::Ident("IF".to_string()),
                Token::LParen,
                Token::CellReference {
                    sheet: None,
                    col_letters: "A".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::Comma,
                Token::Number(1.0),
                Token::Comma,
                Token::Number(2.0),
                Token::RParen,
            ]
        );
    }

    #[test]
    fn named_range() {
        // An identifier not followed by `(` and not matching cell ref pattern
        assert_eq!(toks("MyRange"), vec![Token::Ident("MyRange".to_string())]);
    }

    // ---- Operators ----

    #[test]
    fn all_operators() {
        assert_eq!(
            toks("+ - * / ^ % & = <> < <= > >="),
            vec![
                Token::Plus,
                Token::Minus,
                Token::Star,
                Token::Slash,
                Token::Caret,
                Token::Percent,
                Token::Ampersand,
                Token::Eq,
                Token::Ne,
                Token::Lt,
                Token::Le,
                Token::Gt,
                Token::Ge,
            ]
        );
    }

    #[test]
    fn colon_and_comma() {
        assert_eq!(toks(":,"), vec![Token::Colon, Token::Comma]);
    }

    #[test]
    fn parens() {
        assert_eq!(toks("()"), vec![Token::LParen, Token::RParen]);
    }

    // ---- Leading = ----

    #[test]
    fn leading_equals_stripped() {
        assert_eq!(
            toks("=1+2"),
            vec![Token::Number(1.0), Token::Plus, Token::Number(2.0)]
        );
    }

    // ---- Eof ----

    #[test]
    #[allow(clippy::unwrap_used)]
    fn empty_formula() {
        let tokens = tokenize("").unwrap();
        assert_eq!(tokens, vec![Token::Eof]);
    }

    #[test]
    #[allow(clippy::unwrap_used)]
    fn just_equals() {
        let tokens = tokenize("=").unwrap();
        assert_eq!(tokens, vec![Token::Eof]);
    }

    // ---- Complex formulas ----

    #[test]
    fn complex_formula() {
        let tokens = toks("=SUM(A1:B2, 3.15) + 1");
        assert_eq!(
            tokens,
            vec![
                Token::Ident("SUM".to_string()),
                Token::LParen,
                Token::CellReference {
                    sheet: None,
                    col_letters: "A".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::Colon,
                Token::CellReference {
                    sheet: None,
                    col_letters: "B".to_string(),
                    row: 2,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::Comma,
                Token::Number(3.15),
                Token::RParen,
                Token::Plus,
                Token::Number(1.0),
            ]
        );
    }

    #[test]
    fn nested_function() {
        let tokens = toks("IF(A1>0,SUM(B1:B10),0)");
        assert_eq!(
            tokens,
            vec![
                Token::Ident("IF".to_string()),
                Token::LParen,
                Token::CellReference {
                    sheet: None,
                    col_letters: "A".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::Gt,
                Token::Number(0.0),
                Token::Comma,
                Token::Ident("SUM".to_string()),
                Token::LParen,
                Token::CellReference {
                    sheet: None,
                    col_letters: "B".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::Colon,
                Token::CellReference {
                    sheet: None,
                    col_letters: "B".to_string(),
                    row: 10,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::RParen,
                Token::Comma,
                Token::Number(0.0),
                Token::RParen,
            ]
        );
    }

    #[test]
    fn error_literal_in_formula() {
        let tokens = toks("IF(A1=#N/A,0,A1)");
        assert_eq!(
            tokens,
            vec![
                Token::Ident("IF".to_string()),
                Token::LParen,
                Token::CellReference {
                    sheet: None,
                    col_letters: "A".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::Eq,
                Token::Error(XlError::Na),
                Token::Comma,
                Token::Number(0.0),
                Token::Comma,
                Token::CellReference {
                    sheet: None,
                    col_letters: "A".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::RParen,
            ]
        );
    }

    #[test]
    fn string_concat_formula() {
        let tokens = toks(r#""Hello" & " " & "World""#);
        assert_eq!(
            tokens,
            vec![
                Token::StringLiteral("Hello".to_string()),
                Token::Ampersand,
                Token::StringLiteral(" ".to_string()),
                Token::Ampersand,
                Token::StringLiteral("World".to_string()),
            ]
        );
    }

    #[test]
    fn unrecognized_character() {
        assert!(tokenize("@").is_err());
    }

    #[test]
    fn percent_after_number() {
        assert_eq!(toks("50%"), vec![Token::Number(50.0), Token::Percent]);
    }

    #[test]
    fn comparison_operators() {
        assert_eq!(
            toks("A1<=B1"),
            vec![
                Token::CellReference {
                    sheet: None,
                    col_letters: "A".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::Le,
                Token::CellReference {
                    sheet: None,
                    col_letters: "B".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
            ]
        );
    }

    #[test]
    fn range_with_colon() {
        assert_eq!(
            toks("A1:C5"),
            vec![
                Token::CellReference {
                    sheet: None,
                    col_letters: "A".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::Colon,
                Token::CellReference {
                    sheet: None,
                    col_letters: "C".to_string(),
                    row: 5,
                    col_absolute: false,
                    row_absolute: false,
                },
            ]
        );
    }

    #[test]
    fn multiple_absolute_refs() {
        assert_eq!(
            toks("$A$1:$C$5"),
            vec![
                Token::CellReference {
                    sheet: None,
                    col_letters: "A".to_string(),
                    row: 1,
                    col_absolute: true,
                    row_absolute: true,
                },
                Token::Colon,
                Token::CellReference {
                    sheet: None,
                    col_letters: "C".to_string(),
                    row: 5,
                    col_absolute: true,
                    row_absolute: true,
                },
            ]
        );
    }

    #[test]
    fn function_with_space_before_paren() {
        // SUM (A1) -- space between function name and paren
        assert_eq!(
            toks("SUM (A1)"),
            vec![
                Token::Ident("SUM".to_string()),
                Token::LParen,
                Token::CellReference {
                    sheet: None,
                    col_letters: "A".to_string(),
                    row: 1,
                    col_absolute: false,
                    row_absolute: false,
                },
                Token::RParen,
            ]
        );
    }
}
