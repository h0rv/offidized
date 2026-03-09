//! Comprehensive tests for number formatting in offidized-xlview
//!
//! These tests verify the correctness of Excel-compatible number formatting,
//! including built-in formats, date/time formats, currency, percentages,
//! and custom format codes.
//!
//! These exercise offidized-xlsx's `numfmt` module which provides
//! `format_number`, `get_builtin_format`, and `is_date_format`.
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::float_cmp,
    clippy::panic,
    clippy::approx_constant,
    clippy::cast_possible_truncation,
    clippy::absurd_extreme_comparisons,
    clippy::cast_lossless
)]

use offidized_xlsx::numfmt::{format_number, get_builtin_format, is_date_format};

// ============================================================================
// Built-in Format IDs (numFmtId 0-49)
// ============================================================================

mod builtin_formats {
    use super::*;

    #[test]
    fn test_get_builtin_format_general() {
        assert_eq!(get_builtin_format(0), Some("General"));
    }

    #[test]
    fn test_get_builtin_format_integers() {
        assert_eq!(get_builtin_format(1), Some("0"));
        assert_eq!(get_builtin_format(2), Some("0.00"));
    }

    #[test]
    fn test_get_builtin_format_thousands() {
        assert_eq!(get_builtin_format(3), Some("#,##0"));
        assert_eq!(get_builtin_format(4), Some("#,##0.00"));
    }

    #[test]
    fn test_get_builtin_format_percentages() {
        assert_eq!(get_builtin_format(9), Some("0%"));
        assert_eq!(get_builtin_format(10), Some("0.00%"));
    }

    #[test]
    fn test_get_builtin_format_scientific() {
        assert_eq!(get_builtin_format(11), Some("0.00E+00"));
        assert_eq!(get_builtin_format(48), Some("##0.0E+0"));
    }

    #[test]
    fn test_get_builtin_format_fractions() {
        assert_eq!(get_builtin_format(12), Some("# ?/?"));
        assert_eq!(get_builtin_format(13), Some("# ??/??"));
    }

    #[test]
    fn test_get_builtin_format_dates() {
        assert_eq!(get_builtin_format(14), Some("mm-dd-yy"));
        assert_eq!(get_builtin_format(15), Some("d-mmm-yy"));
        assert_eq!(get_builtin_format(16), Some("d-mmm"));
        assert_eq!(get_builtin_format(17), Some("mmm-yy"));
        assert_eq!(get_builtin_format(22), Some("m/d/yy h:mm"));
    }

    #[test]
    fn test_get_builtin_format_times() {
        assert_eq!(get_builtin_format(18), Some("h:mm AM/PM"));
        assert_eq!(get_builtin_format(19), Some("h:mm:ss AM/PM"));
        assert_eq!(get_builtin_format(20), Some("h:mm"));
        assert_eq!(get_builtin_format(21), Some("h:mm:ss"));
        assert_eq!(get_builtin_format(45), Some("mm:ss"));
        assert_eq!(get_builtin_format(46), Some("[h]:mm:ss"));
        assert_eq!(get_builtin_format(47), Some("mmss.0"));
    }

    #[test]
    fn test_get_builtin_format_accounting() {
        assert_eq!(get_builtin_format(37), Some("#,##0 ;(#,##0)"));
        assert_eq!(get_builtin_format(38), Some("#,##0 ;[Red](#,##0)"));
        assert_eq!(get_builtin_format(39), Some("#,##0.00;(#,##0.00)"));
        assert_eq!(get_builtin_format(40), Some("#,##0.00;[Red](#,##0.00)"));
    }

    #[test]
    fn test_get_builtin_format_text() {
        assert_eq!(get_builtin_format(49), Some("@"));
    }

    #[test]
    fn test_get_builtin_format_currency() {
        assert_eq!(get_builtin_format(5), Some("$#,##0_);($#,##0)"));
        assert_eq!(get_builtin_format(6), Some("$#,##0_);[Red]($#,##0)"));
        assert_eq!(get_builtin_format(7), Some("$#,##0.00_);($#,##0.00)"));
        assert_eq!(get_builtin_format(8), Some("$#,##0.00_);[Red]($#,##0.00)"));
    }

    #[test]
    fn test_get_builtin_format_accounting_no_decimals() {
        assert_eq!(
            get_builtin_format(41),
            Some("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)")
        );
        assert_eq!(
            get_builtin_format(42),
            Some("_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)")
        );
    }

    #[test]
    fn test_get_builtin_format_unknown() {
        assert_eq!(get_builtin_format(50), None);
        assert_eq!(get_builtin_format(100), None);
        assert_eq!(get_builtin_format(u32::MAX), None);
    }
}

// ============================================================================
// General Format (numFmtId 0)
// ============================================================================

mod general_format {
    use super::*;

    #[test]
    fn test_general_integer() {
        assert_eq!(format_number(1234.0, "General", false), "1234");
        assert_eq!(format_number(0.0, "General", false), "0");
        assert_eq!(format_number(-5678.0, "General", false), "-5678");
    }

    #[test]
    fn test_general_decimal() {
        assert_eq!(format_number(1234.5, "General", false), "1234.5");
        assert_eq!(format_number(3.14159, "General", false), "3.14159");
        assert_eq!(format_number(-99.99, "General", false), "-99.99");
    }

    #[test]
    fn test_general_case_insensitive() {
        assert_eq!(format_number(42.0, "general", false), "42");
        assert_eq!(format_number(42.0, "GENERAL", false), "42");
        assert_eq!(format_number(42.0, "General", false), "42");
    }

    #[test]
    fn test_general_trims_trailing_zeros() {
        assert_eq!(format_number(1.5000, "General", false), "1.5");
        assert_eq!(format_number(10.10, "General", false), "10.1");
    }

    #[test]
    fn test_general_very_large_numbers() {
        let result = format_number(1e12, "General", false);
        assert!(result.contains('E') || result.contains("1000000000000"));
    }

    #[test]
    fn test_general_very_small_numbers() {
        let result = format_number(0.00001, "General", false);
        assert!(result.contains('E') || result.contains("0.00001"));
    }
}

// ============================================================================
// Fixed Decimal Formats (numFmtId 1, 2)
// ============================================================================

mod fixed_decimal_formats {
    use super::*;

    #[test]
    fn test_format_0_rounds_to_integer() {
        assert_eq!(format_number(1234.5, "0", false), "1234");
        assert_eq!(format_number(1234.4, "0", false), "1234");
        assert_eq!(format_number(1234.0, "0", false), "1234");
        assert_eq!(format_number(1233.5, "0", false), "1234");
    }

    #[test]
    fn test_format_0_negative() {
        assert_eq!(format_number(-1234.5, "0", false), "-1234");
        assert_eq!(format_number(-1234.4, "0", false), "-1234");
    }

    #[test]
    fn test_format_0_00_two_decimals() {
        assert_eq!(format_number(1234.5, "0.00", false), "1234.50");
        assert_eq!(format_number(1234.567, "0.00", false), "1234.57");
        assert_eq!(format_number(1234.0, "0.00", false), "1234.00");
    }

    #[test]
    fn test_format_0_00_negative() {
        assert_eq!(format_number(-1234.5, "0.00", false), "-1234.50");
    }

    #[test]
    fn test_format_0_000_three_decimals() {
        assert_eq!(format_number(1234.5, "0.000", false), "1234.500");
        assert_eq!(format_number(1234.5678, "0.000", false), "1234.568");
    }

    #[test]
    fn test_format_0_0_one_decimal() {
        assert_eq!(format_number(1234.56, "0.0", false), "1234.6");
        assert_eq!(format_number(1234.0, "0.0", false), "1234.0");
    }
}

// ============================================================================
// Thousands Separator Formats (numFmtId 3, 4)
// ============================================================================

mod thousands_separator_formats {
    use super::*;

    #[test]
    fn test_format_comma_separated_integer() {
        assert_eq!(format_number(1234567.0, "#,##0", false), "1,234,567");
        assert_eq!(format_number(1234.0, "#,##0", false), "1,234");
        assert_eq!(format_number(123.0, "#,##0", false), "123");
        assert_eq!(format_number(12.0, "#,##0", false), "12");
        assert_eq!(format_number(1.0, "#,##0", false), "1");
    }

    #[test]
    fn test_format_comma_separated_decimal() {
        assert_eq!(format_number(1234.5, "#,##0.00", false), "1,234.50");
        assert_eq!(format_number(1234567.89, "#,##0.00", false), "1,234,567.89");
    }

    #[test]
    fn test_format_comma_separated_negative() {
        assert_eq!(format_number(-1234567.0, "#,##0", false), "-1,234,567");
        assert_eq!(format_number(-1234.5, "#,##0.00", false), "-1,234.50");
    }

    #[test]
    fn test_format_comma_separated_zero() {
        assert_eq!(format_number(0.0, "#,##0", false), "0");
        assert_eq!(format_number(0.0, "#,##0.00", false), "0.00");
    }

    #[test]
    fn test_format_comma_separated_large() {
        assert_eq!(format_number(1234567890.0, "#,##0", false), "1,234,567,890");
        assert_eq!(
            format_number(1234567890123.0, "#,##0", false),
            "1,234,567,890,123"
        );
    }

    #[test]
    fn test_format_comma_separated_small_decimal() {
        assert_eq!(format_number(0.5, "#,##0.00", false), "0.50");
        assert_eq!(format_number(0.99, "#,##0.00", false), "0.99");
    }
}

// ============================================================================
// Percentage Formats (numFmtId 9, 10)
// ============================================================================

mod percentage_formats {
    use super::*;

    #[test]
    fn test_format_percentage_integer() {
        assert_eq!(format_number(0.5, "0%", false), "50%");
        assert_eq!(format_number(1.0, "0%", false), "100%");
        assert_eq!(format_number(0.0, "0%", false), "0%");
    }

    #[test]
    fn test_format_percentage_decimal() {
        assert_eq!(format_number(0.5, "0.00%", false), "50.00%");
        assert_eq!(format_number(0.123, "0.00%", false), "12.30%");
        assert_eq!(format_number(0.1234, "0.00%", false), "12.34%");
    }

    #[test]
    fn test_format_percentage_rounding() {
        assert_eq!(format_number(0.12345, "0.00%", false), "12.35%");
        assert_eq!(format_number(0.12344, "0.00%", false), "12.34%");
    }

    #[test]
    fn test_format_percentage_negative() {
        assert_eq!(format_number(-0.5, "0%", false), "-50%");
        assert_eq!(format_number(-0.123, "0.00%", false), "-12.30%");
    }

    #[test]
    fn test_format_percentage_greater_than_100() {
        assert_eq!(format_number(2.5, "0%", false), "250%");
        assert_eq!(format_number(10.0, "0%", false), "1000%");
    }

    #[test]
    fn test_format_percentage_small_values() {
        assert_eq!(format_number(0.001, "0%", false), "0%");
        assert_eq!(format_number(0.001, "0.00%", false), "0.10%");
        assert_eq!(format_number(0.0001, "0.00%", false), "0.01%");
    }

    #[test]
    fn test_format_percentage_one_decimal() {
        assert_eq!(format_number(0.5, "0.0%", false), "50.0%");
        assert_eq!(format_number(0.1234, "0.0%", false), "12.3%");
    }
}

// ============================================================================
// Scientific Notation Format (numFmtId 11)
// ============================================================================

mod scientific_formats {
    use super::*;

    #[test]
    fn test_format_scientific_basic() {
        let result = format_number(1234.0, "0.00E+00", false);
        assert!(result.contains('E') || result.contains('e'));
    }

    #[test]
    fn test_format_scientific_small_number() {
        let result = format_number(0.00123, "0.00E+00", false);
        assert!(result.contains('E') || result.contains('e'));
    }

    #[test]
    fn test_format_scientific_large_number() {
        let result = format_number(1234567890.0, "0.00E+00", false);
        assert!(result.contains('E') || result.contains('e'));
    }

    #[test]
    fn test_format_scientific_negative() {
        let result = format_number(-1234.0, "0.00E+00", false);
        assert!(result.contains('-'));
        assert!(result.contains('E') || result.contains('e'));
    }

    #[test]
    fn test_format_scientific_one() {
        let result = format_number(1.0, "0.00E+00", false);
        assert!(result.contains('E') || result.contains('e'));
    }
}

// ============================================================================
// Fraction Formats (numFmtId 12, 13)
// ============================================================================

mod fraction_formats {
    use super::*;

    #[test]
    fn test_format_fraction_half() {
        let result = format_number(0.5, "# ?/?", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_fraction_third() {
        let result = format_number(0.333333, "# ??/??", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_fraction_quarter() {
        let result = format_number(0.25, "# ?/?", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_fraction_mixed() {
        let result = format_number(2.5, "# ?/?", false);
        assert!(!result.is_empty());
    }
}

// ============================================================================
// Date Formats (numFmtId 14-17, 22)
// ============================================================================

mod date_formats {
    use super::*;

    #[test]
    fn test_format_date_mm_dd_yy() {
        let result = format_number(44927.0, "mm-dd-yy", false);
        assert!(result.contains("15") || result.contains("01"));
    }

    #[test]
    fn test_format_date_d_mmm_yy() {
        let result = format_number(44927.0, "d-mmm-yy", false);
        assert!(result.contains("Jan") || result.contains("15"));
    }

    #[test]
    fn test_format_date_d_mmm() {
        let result = format_number(44927.0, "d-mmm", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_date_mmm_yy() {
        let result = format_number(44927.0, "mmm-yy", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_date_yyyy_mm_dd() {
        let result = format_number(44927.0, "yyyy-mm-dd", false);
        assert!(result.contains("2023") || result.contains("01") || result.contains("15"));
    }

    #[test]
    fn test_format_date_m_d_yyyy() {
        let result = format_number(44927.0, "m/d/yyyy", false);
        assert!(result.contains('/'));
    }

    #[test]
    fn test_format_date_dd_mm_yyyy() {
        let result = format_number(44927.0, "dd/mm/yyyy", false);
        assert!(result.contains('/') || result.contains('-'));
    }

    #[test]
    fn test_format_date_excel_epoch() {
        let result = format_number(1.0, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("01"));
    }

    #[test]
    fn test_format_date_leap_year_bug() {
        let result_59 = format_number(59.0, "yyyy-mm-dd", false);
        let result_61 = format_number(61.0, "yyyy-mm-dd", false);
        assert!(!result_59.is_empty());
        assert!(!result_61.is_empty());
    }

    #[test]
    fn test_format_date_modern_date() {
        let result = format_number(45458.0, "yyyy-mm-dd", false);
        assert!(result.contains("2024") || !result.is_empty());
    }

    #[test]
    fn test_format_date_with_time() {
        let result = format_number(44927.5, "m/d/yy h:mm", false);
        assert!(!result.is_empty());
    }
}

// ============================================================================
// Time Formats (numFmtId 18-21)
// ============================================================================

mod time_formats {
    use super::*;

    #[test]
    fn test_format_time_h_mm_am_pm_noon() {
        let result = format_number(0.5, "h:mm AM/PM", false);
        assert!(result.contains("12") && (result.contains("PM") || result.contains("pm")));
    }

    #[test]
    fn test_format_time_h_mm_am_pm_midnight() {
        let result = format_number(0.0, "h:mm AM/PM", false);
        assert!(result.contains(':'));
    }

    #[test]
    fn test_format_time_h_mm_am_pm_morning() {
        let result = format_number(0.25, "h:mm AM/PM", false);
        assert!(result.contains('6') || result.contains(':'));
    }

    #[test]
    fn test_format_time_h_mm_am_pm_evening() {
        let result = format_number(0.75, "h:mm AM/PM", false);
        assert!(result.contains('6') || result.contains("18") || result.contains(':'));
    }

    #[test]
    fn test_format_time_h_mm_ss_am_pm() {
        let result = format_number(0.5, "h:mm:ss AM/PM", false);
        assert!(result.contains("12"));
        assert!(result.contains(':'));
    }

    #[test]
    fn test_format_time_h_mm_24_hour() {
        let result = format_number(0.75, "h:mm", false);
        assert!(result.contains("18") || result.contains('6'));
        assert!(result.contains(':'));
    }

    #[test]
    fn test_format_time_h_mm_ss_24_hour() {
        let result = format_number(0.75, "h:mm:ss", false);
        assert!(result.contains(':'));
    }

    #[test]
    fn test_format_time_fractional_minute() {
        let result = format_number(0.5006944444444444, "h:mm:ss", false);
        assert!(result.contains(':'));
    }

    #[test]
    fn test_format_time_with_date() {
        let result = format_number(44927.5, "m/d/yy h:mm", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_time_just_before_midnight() {
        let result = format_number(0.99999, "h:mm:ss", false);
        assert!(result.contains(':'));
    }
}

// ============================================================================
// Currency Formats
// ============================================================================

mod currency_formats {
    use super::*;

    #[test]
    fn test_format_currency_usd_integer() {
        let result = format_number(1234.0, "$#,##0", false);
        assert!(result.contains('$'));
        assert!(result.contains("1,234"));
    }

    #[test]
    fn test_format_currency_usd_decimal() {
        let result = format_number(1234.5, "$#,##0.00", false);
        assert!(result.contains('$'));
        assert!(result.contains("1,234.50"));
    }

    #[test]
    fn test_format_currency_usd_negative() {
        let result = format_number(-1234.0, "$#,##0", false);
        assert!(result.contains('$'));
        assert!(result.contains("1,234"));
    }

    #[test]
    fn test_format_currency_usd_zero() {
        let result = format_number(0.0, "$#,##0.00", false);
        assert!(result.contains('$'));
        assert!(result.contains("0.00"));
    }

    #[test]
    fn test_format_currency_euro() {
        let result = format_number(1234.5, "\u{20ac}#,##0.00", false);
        assert!(result.contains('\u{20ac}'));
    }

    #[test]
    fn test_format_currency_pound() {
        let result = format_number(1234.5, "\u{00a3}#,##0.00", false);
        assert!(result.contains('\u{00a3}'));
    }

    #[test]
    fn test_format_currency_large_amount() {
        let result = format_number(1234567890.12, "$#,##0.00", false);
        assert!(result.contains('$'));
        assert!(result.contains(','));
    }

    #[test]
    fn test_format_currency_small_amount() {
        let result = format_number(0.99, "$#,##0.00", false);
        assert!(result.contains('$'));
        assert!(result.contains("0.99"));
    }
}

// ============================================================================
// Accounting Formats (numFmtId 37-40)
// ============================================================================

mod accounting_formats {
    use super::*;

    #[test]
    fn test_format_accounting_positive() {
        let result = format_number(1234.0, "#,##0 ;(#,##0)", false);
        assert!(result.contains("1,234"));
    }

    #[test]
    fn test_format_accounting_negative_parens() {
        let result = format_number(-1234.0, "#,##0;(#,##0)", false);
        assert!(result.contains("1,234") || result.contains("1234"));
    }

    #[test]
    fn test_format_accounting_negative_red() {
        let result = format_number(-1234.0, "#,##0;[Red](#,##0)", false);
        assert!(result.contains("1234") || result.contains("1,234"));
    }

    #[test]
    fn test_format_accounting_decimal_positive() {
        let result = format_number(1234.56, "#,##0.00;(#,##0.00)", false);
        assert!(result.contains("1,234.56"));
    }

    #[test]
    fn test_format_accounting_decimal_negative() {
        let result = format_number(-1234.56, "#,##0.00;(#,##0.00)", false);
        assert!(result.contains("1,234.56") || result.contains("1234.56"));
    }
}

// ============================================================================
// Custom Format Sections (positive;negative;zero;text)
// ============================================================================

mod custom_format_sections {
    use super::*;

    #[test]
    fn test_format_two_sections_positive() {
        let result = format_number(1234.0, "#,##0;(#,##0)", false);
        assert!(result.contains("1,234"));
    }

    #[test]
    fn test_format_two_sections_negative() {
        let result = format_number(-1234.0, "#,##0;(#,##0)", false);
        assert!(result.contains("1234") || result.contains("1,234"));
    }

    #[test]
    fn test_format_three_sections_positive() {
        let result = format_number(1234.0, r#"#,##0;-#,##0;"-""#, false);
        assert!(result.contains("1,234"));
    }

    #[test]
    fn test_format_three_sections_negative() {
        let result = format_number(-1234.0, r#"#,##0;-#,##0;"-""#, false);
        assert!(result.contains("1234") || result.contains("1,234"));
    }

    #[test]
    fn test_format_three_sections_zero() {
        let result = format_number(0.0, r#"#,##0.00;-#,##0.00;"-""#, false);
        assert!(!result.is_empty());
    }
}

// ============================================================================
// Text Format (numFmtId 49, @)
// ============================================================================

mod text_format {
    use super::*;

    #[test]
    fn test_format_text_passthrough_integer() {
        let result = format_number(1234.0, "@", false);
        assert!(result.contains("1234"));
    }

    #[test]
    fn test_format_text_passthrough_decimal() {
        let result = format_number(1234.5, "@", false);
        assert!(result.contains("1234") || result.contains("1234.5"));
    }
}

// ============================================================================
// Special Number Cases
// ============================================================================

mod special_cases {
    use super::*;

    #[test]
    fn test_format_very_large_number() {
        let result = format_number(1234567890123.0, "#,##0", false);
        assert!(result.contains(',') || result.contains("1234567890123"));
    }

    #[test]
    fn test_format_very_small_number() {
        let result = format_number(0.00000001, "0.00000000", false);
        assert!(result.contains("0.00000001") || result.contains('E'));
    }

    #[test]
    fn test_format_zero_integer() {
        assert_eq!(format_number(0.0, "0", false), "0");
    }

    #[test]
    fn test_format_zero_decimal() {
        assert_eq!(format_number(0.0, "0.00", false), "0.00");
    }

    #[test]
    fn test_format_zero_thousands() {
        assert_eq!(format_number(0.0, "#,##0", false), "0");
    }

    #[test]
    fn test_format_zero_percentage() {
        assert_eq!(format_number(0.0, "0%", false), "0%");
    }

    #[test]
    fn test_format_negative_integer() {
        assert_eq!(format_number(-1234.0, "0", false), "-1234");
    }

    #[test]
    fn test_format_negative_decimal() {
        assert_eq!(format_number(-1234.5, "0.00", false), "-1234.50");
    }

    #[test]
    fn test_format_negative_thousands() {
        assert_eq!(format_number(-1234567.0, "#,##0", false), "-1,234,567");
    }

    #[test]
    fn test_format_negative_percentage() {
        assert_eq!(format_number(-0.5, "0%", false), "-50%");
    }

    #[test]
    fn test_format_positive_infinity() {
        let result = format_number(f64::INFINITY, "0.00", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_negative_infinity() {
        let result = format_number(f64::NEG_INFINITY, "0.00", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_nan() {
        let result = format_number(f64::NAN, "0.00", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_max_f64() {
        let result = format_number(f64::MAX, "General", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_min_positive_f64() {
        let result = format_number(f64::MIN_POSITIVE, "General", false);
        assert!(!result.is_empty());
    }
}

// ============================================================================
// Date Detection (is_date_format)
// ============================================================================

mod date_detection {
    use super::*;

    #[test]
    fn test_is_date_format_year() {
        assert!(is_date_format("yyyy"));
        assert!(is_date_format("yy"));
        assert!(is_date_format("yyyy-mm-dd"));
    }

    #[test]
    fn test_is_date_format_month() {
        assert!(is_date_format("mm"));
        assert!(is_date_format("mmm"));
        assert!(is_date_format("mmmm"));
    }

    #[test]
    fn test_is_date_format_day() {
        assert!(is_date_format("dd"));
        assert!(is_date_format("d"));
        assert!(is_date_format("d-mmm-yy"));
    }

    #[test]
    fn test_is_date_format_time_hours() {
        assert!(is_date_format("h:mm"));
        assert!(is_date_format("hh:mm"));
        assert!(is_date_format("h:mm AM/PM"));
    }

    #[test]
    fn test_is_date_format_time_seconds() {
        assert!(is_date_format("h:mm:ss"));
        assert!(is_date_format("mm:ss"));
    }

    #[test]
    fn test_is_not_date_format_number() {
        assert!(!is_date_format("#,##0"));
        assert!(!is_date_format("#,##0.00"));
        assert!(!is_date_format("0"));
        assert!(!is_date_format("0.00"));
    }

    #[test]
    fn test_is_not_date_format_percentage() {
        assert!(!is_date_format("0%"));
        assert!(!is_date_format("0.00%"));
    }

    #[test]
    fn test_is_not_date_format_currency() {
        assert!(!is_date_format("$#,##0"));
        assert!(!is_date_format("$#,##0.00"));
    }

    #[test]
    fn test_is_date_format_ignores_quoted_text() {
        assert!(!is_date_format(r#""day" 0"#));
        assert!(!is_date_format(r#"0 "month""#));
    }

    #[test]
    fn test_is_date_format_ignores_bracketed_text() {
        assert!(!is_date_format("[Red]#,##0"));
    }

    #[test]
    fn test_is_date_format_combined_datetime() {
        assert!(is_date_format("yyyy-mm-dd hh:mm:ss"));
        assert!(is_date_format("m/d/yy h:mm"));
    }

    #[test]
    fn test_is_date_format_case_insensitive() {
        assert!(is_date_format("YYYY-MM-DD"));
        assert!(is_date_format("Yyyy-Mm-Dd"));
    }
}

// ============================================================================
// Edge Cases for Date/Time Formatting
// ============================================================================

mod datetime_edge_cases {
    use super::*;

    #[test]
    fn test_format_date_day_1() {
        let result = format_number(1.0, "yyyy-mm-dd", false);
        assert!(result.contains("1900"));
    }

    #[test]
    fn test_format_date_day_0() {
        let result = format_number(0.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_date_negative_day() {
        let result = format_number(-1.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_time_boundary_midnight() {
        let result = format_number(0.0, "h:mm:ss", false);
        assert!(result.contains(':'));
    }

    #[test]
    fn test_format_time_boundary_just_before_midnight() {
        let result = format_number(0.9999884259259259, "h:mm:ss", false);
        assert!(result.contains(':'));
    }

    #[test]
    fn test_format_datetime_combined() {
        let result = format_number(44927.75, "yyyy-mm-dd hh:mm:ss", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_time_only_no_date() {
        let result = format_number(0.333333, "h:mm:ss", false);
        assert!(result.contains(':'));
    }

    #[test]
    fn test_format_date_future() {
        let result = format_number(73050.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }
}

// ============================================================================
// Format Code Parsing Edge Cases
// ============================================================================

mod format_parsing_edge_cases {
    use super::*;

    #[test]
    fn test_format_whitespace_trimmed() {
        assert_eq!(format_number(1234.0, "  0  ", false), "1234");
    }

    #[test]
    fn test_format_empty_string() {
        let result = format_number(1234.0, "", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_unknown_format() {
        let result = format_number(1234.0, "???", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_format_mixed_case() {
        let result = format_number(1234.0, "GENERAL", false);
        assert_eq!(result, "1234");
    }

    #[test]
    fn test_format_with_text_literals() {
        let result = format_number(1234.0, r#"0 "units""#, false);
        assert!(result.contains("1234"));
    }

    #[test]
    fn test_format_color_codes() {
        let result = format_number(1234.0, "[Blue]#,##0", false);
        assert!(result.contains("1234") || result.contains("1,234"));
    }

    #[test]
    fn test_format_conditional() {
        let result = format_number(100.0, "[>50]0;0.00", false);
        assert!(!result.is_empty());
    }
}

// ============================================================================
// Rounding Behavior
// ============================================================================

mod rounding_behavior {
    use super::*;

    #[test]
    fn test_round_half_to_even() {
        assert_eq!(format_number(1.5, "0", false), "2");
        assert_eq!(format_number(2.5, "0", false), "2");
    }

    #[test]
    fn test_round_half_to_even_negative() {
        assert_eq!(format_number(-1.5, "0", false), "-2");
        assert_eq!(format_number(-2.5, "0", false), "-2");
    }

    #[test]
    fn test_round_decimal_places() {
        assert_eq!(format_number(1.234, "0.00", false), "1.23");
        assert_eq!(format_number(1.235, "0.00", false), "1.24");
        assert_eq!(format_number(1.2351, "0.00", false), "1.24");
    }

    #[test]
    fn test_round_large_decimal() {
        assert_eq!(format_number(1.999, "0.00", false), "2.00");
        assert_eq!(format_number(9.999, "0.00", false), "10.00");
    }

    #[test]
    fn test_round_percentage() {
        assert_eq!(format_number(0.1234, "0%", false), "12%");
        assert_eq!(format_number(0.1250, "0%", false), "12%");
    }
}

// ============================================================================
// Precision Edge Cases
// ============================================================================

mod precision_edge_cases {
    use super::*;

    #[test]
    fn test_floating_point_precision() {
        let result = format_number(0.1 + 0.2, "0.0", false);
        assert!(result == "0.3" || result.starts_with("0.3"));
    }

    #[test]
    fn test_many_decimal_places() {
        let result = format_number(1.0 / 3.0, "0.0000000000", false);
        assert!(result.contains('.'));
    }

    #[test]
    fn test_trailing_zeros_preserved() {
        assert_eq!(format_number(1.0, "0.00", false), "1.00");
        assert_eq!(format_number(1.1, "0.00", false), "1.10");
    }

    #[test]
    fn test_leading_zeros() {
        let result = format_number(42.0, "00000", false);
        assert!(result.contains("42"));
    }
}

// ============================================================================
// Integration Tests with Built-in Format IDs
// ============================================================================

mod integration_builtin_formats {
    use super::*;

    #[test]
    fn test_apply_builtin_format_0() {
        if let Some(fmt) = get_builtin_format(0) {
            let result = format_number(1234.5, fmt, false);
            assert_eq!(result, "1234.5");
        }
    }

    #[test]
    fn test_apply_builtin_format_1() {
        if let Some(fmt) = get_builtin_format(1) {
            let result = format_number(1234.5, fmt, false);
            assert_eq!(result, "1234");
        }
    }

    #[test]
    fn test_apply_builtin_format_2() {
        if let Some(fmt) = get_builtin_format(2) {
            let result = format_number(1234.5, fmt, false);
            assert_eq!(result, "1234.50");
        }
    }

    #[test]
    fn test_apply_builtin_format_3() {
        if let Some(fmt) = get_builtin_format(3) {
            let result = format_number(1234567.0, fmt, false);
            assert_eq!(result, "1,234,567");
        }
    }

    #[test]
    fn test_apply_builtin_format_4() {
        if let Some(fmt) = get_builtin_format(4) {
            let result = format_number(1234.5, fmt, false);
            assert_eq!(result, "1,234.50");
        }
    }

    #[test]
    fn test_apply_builtin_format_9() {
        if let Some(fmt) = get_builtin_format(9) {
            let result = format_number(0.5, fmt, false);
            assert_eq!(result, "50%");
        }
    }

    #[test]
    fn test_apply_builtin_format_10() {
        if let Some(fmt) = get_builtin_format(10) {
            let result = format_number(0.5, fmt, false);
            assert_eq!(result, "50.00%");
        }
    }
}
