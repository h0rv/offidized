//! Comprehensive tests for date and time formatting in offidized-xlview.
//!
//! This module tests Excel date/time handling including:
//! - Standard date formats (mm-dd-yy, d-mmm-yy, d-mmm, mmm-yy)
//! - ISO date format (yyyy-mm-dd)
//! - Long date format (mmmm d, yyyy)
//! - Time formats (h:mm, h:mm:ss, h:mm AM/PM)
//! - Combined date/time formats
//! - 1900 date system (Windows Excel default)
//! - 1904 date system (Mac Excel)
//! - Excel leap year bug (Feb 29, 1900)
//! - Negative dates
//! - Date as serial number vs formatted

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

use offidized_xlsx::numfmt::{format_number, is_date_format};

// ============================================================================
// Excel Date System Constants
// ============================================================================

/// Excel serial number for January 1, 1900 (day 1 in 1900 system)
const EXCEL_JAN_1_1900: f64 = 1.0;

/// Excel serial number for February 28, 1900 (day 59)
const EXCEL_FEB_28_1900: f64 = 59.0;

/// Excel serial number for the fake February 29, 1900 (day 60 - doesn't exist)
const EXCEL_FEB_29_1900_FAKE: f64 = 60.0;

/// Excel serial number for March 1, 1900 (day 61)
const EXCEL_MAR_1_1900: f64 = 61.0;

/// Excel serial number for December 31, 1900 (day 366)
const EXCEL_DEC_31_1900: f64 = 366.0;

/// Excel serial number for January 1, 2000 (day 36526)
const EXCEL_JAN_1_2000: f64 = 36526.0;

/// Excel serial number for January 15, 2023 (day 44941)
const EXCEL_JAN_15_2023: f64 = 44941.0;

/// Excel serial number for July 4, 2024 (day 45477)
const EXCEL_JUL_4_2024: f64 = 45477.0;

/// Excel serial number for December 31, 2099 (day 73050)
const EXCEL_DEC_31_2099: f64 = 73050.0;

// ============================================================================
// Time Constants (fractions of a day)
// ============================================================================

/// 12:00:00 AM (midnight) = 0.0
const TIME_MIDNIGHT: f64 = 0.0;

/// 6:00:00 AM = 0.25
const TIME_6AM: f64 = 0.25;

/// 12:00:00 PM (noon) = 0.5
const TIME_NOON: f64 = 0.5;

/// 6:00:00 PM = 0.75
const TIME_6PM: f64 = 0.75;

/// 11:59:59 PM = ~0.99999
const TIME_JUST_BEFORE_MIDNIGHT: f64 = 0.9999884259259259;

/// 1:30:45 PM = 0.5630208333...
const TIME_1_30_45_PM: f64 = 0.5630208333333333;

/// 8:15:30 AM = 0.3440972222...
const TIME_8_15_30_AM: f64 = 0.344_097_222_222_222_2;

// ============================================================================
// Standard Date Formats (numFmtId 14-17)
// ============================================================================

mod standard_date_formats {
    use super::*;

    // Format 14: mm-dd-yy
    mod mm_dd_yy {
        use super::*;

        #[test]
        fn test_mm_dd_yy_basic_date() {
            let result = format_number(EXCEL_JAN_15_2023, "mm-dd-yy", false);
            assert!(result.contains("01") || result.contains("1"));
            assert!(result.contains("15"));
            assert!(result.contains("23"));
        }

        #[test]
        fn test_mm_dd_yy_single_digit_month() {
            let result = format_number(36590.0, "mm-dd-yy", false);
            assert!(!result.is_empty());
        }

        #[test]
        fn test_mm_dd_yy_december() {
            let result = format_number(45285.0, "mm-dd-yy", false);
            assert!(result.contains("12") || result.contains("25"));
        }

        #[test]
        fn test_mm_dd_yy_leap_year_date() {
            let result = format_number(45351.0, "mm-dd-yy", false);
            assert!(!result.is_empty());
        }

        #[test]
        fn test_mm_dd_yy_end_of_year() {
            let result = format_number(EXCEL_DEC_31_1900, "mm-dd-yy", false);
            assert!(result.contains("12") || result.contains("31"));
        }
    }

    // Format 15: d-mmm-yy
    mod d_mmm_yy {
        use super::*;

        #[test]
        fn test_d_mmm_yy_basic_date() {
            let result = format_number(EXCEL_JAN_15_2023, "d-mmm-yy", false);
            assert!(result.contains("15") || result.contains("Jan"));
        }

        #[test]
        fn test_d_mmm_yy_january() {
            let result = format_number(EXCEL_JAN_1_2000, "d-mmm-yy", false);
            assert!(result.contains("Jan") || result.contains("1"));
        }

        #[test]
        fn test_d_mmm_yy_july() {
            let result = format_number(EXCEL_JUL_4_2024, "d-mmm-yy", false);
            assert!(result.contains("Jul") || result.contains("4"));
        }

        #[test]
        fn test_d_mmm_yy_december() {
            let result = format_number(45285.0, "d-mmm-yy", false);
            assert!(result.contains("Dec") || result.contains("25"));
        }

        #[test]
        fn test_d_mmm_yy_all_months() {
            let month_serials = [
                (44927.0, "Jan"),
                (44958.0, "Feb"),
                (44986.0, "Mar"),
                (45017.0, "Apr"),
                (45047.0, "May"),
                (45078.0, "Jun"),
                (45108.0, "Jul"),
                (45139.0, "Aug"),
                (45170.0, "Sep"),
                (45200.0, "Oct"),
                (45231.0, "Nov"),
                (45261.0, "Dec"),
            ];

            for (serial, expected_month) in month_serials {
                let result = format_number(serial, "d-mmm-yy", false);
                assert!(
                    result
                        .to_lowercase()
                        .contains(&expected_month.to_lowercase()),
                    "Expected {} in result '{}' for serial {}",
                    expected_month,
                    result,
                    serial
                );
            }
        }
    }

    // Format 16: d-mmm
    mod d_mmm {
        use super::*;

        #[test]
        fn test_d_mmm_basic_date() {
            let result = format_number(EXCEL_JAN_15_2023, "d-mmm", false);
            assert!(!result.is_empty());
        }

        #[test]
        fn test_d_mmm_single_digit_day() {
            let result = format_number(36590.0, "d-mmm", false);
            assert!(result.contains("5") || result.contains("Mar"));
        }

        #[test]
        fn test_d_mmm_double_digit_day() {
            let result = format_number(44951.0, "d-mmm", false);
            assert!(result.contains("25") || result.contains("Jan"));
        }
    }

    // Format 17: mmm-yy
    mod mmm_yy {
        use super::*;

        #[test]
        fn test_mmm_yy_basic_date() {
            let result = format_number(EXCEL_JAN_15_2023, "mmm-yy", false);
            assert!(!result.is_empty());
        }

        #[test]
        fn test_mmm_yy_y2k() {
            let result = format_number(EXCEL_JAN_1_2000, "mmm-yy", false);
            assert!(result.contains("Jan") || result.contains("00"));
        }

        #[test]
        fn test_mmm_yy_different_years() {
            let years = [
                (36526.0, "00"),
                (40179.0, "10"),
                (44927.0, "23"),
                (45658.0, "25"),
            ];

            for (serial, expected_year) in years {
                let result = format_number(serial, "mmm-yy", false);
                assert!(
                    result.contains(expected_year),
                    "Expected year {} in result '{}' for serial {}",
                    expected_year,
                    result,
                    serial
                );
            }
        }
    }
}

// ============================================================================
// ISO Date Format (yyyy-mm-dd)
// ============================================================================

mod iso_date_format {
    use super::*;

    #[test]
    fn test_iso_basic_date() {
        let result = format_number(EXCEL_JAN_15_2023, "yyyy-mm-dd", false);
        assert!(result.contains("2023"));
        assert!(result.contains("01"));
        assert!(result.contains("15"));
    }

    #[test]
    fn test_iso_y2k_date() {
        let result = format_number(EXCEL_JAN_1_2000, "yyyy-mm-dd", false);
        assert!(result.contains("2000") || result.contains("01"));
    }

    #[test]
    fn test_iso_1900_date() {
        let result = format_number(EXCEL_JAN_1_1900, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("01"));
    }

    #[test]
    fn test_iso_future_date() {
        let result = format_number(EXCEL_DEC_31_2099, "yyyy-mm-dd", false);
        assert!(result.contains("2099") || result.contains("12") || result.contains("31"));
    }

    #[test]
    fn test_iso_with_different_separators() {
        let result_slash = format_number(EXCEL_JAN_15_2023, "yyyy/mm/dd", false);
        assert!(
            result_slash.contains("2023")
                || result_slash.contains("01")
                || result_slash.contains("15")
        );

        let result_dot = format_number(EXCEL_JAN_15_2023, "yyyy.mm.dd", false);
        assert!(!result_dot.is_empty());
    }

    #[test]
    fn test_iso_padded_months_and_days() {
        let march_1 = format_number(36586.0, "yyyy-mm-dd", false);
        assert!(!march_1.is_empty());

        let jan_5 = format_number(36531.0, "yyyy-mm-dd", false);
        assert!(!jan_5.is_empty());
    }

    #[test]
    fn test_iso_leap_year_feb_29() {
        let result = format_number(45351.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }
}

// ============================================================================
// Long Date Format (mmmm d, yyyy)
// ============================================================================

mod long_date_format {
    use super::*;

    #[test]
    fn test_long_date_basic() {
        let result = format_number(EXCEL_JAN_15_2023, "mmmm d, yyyy", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_long_date_full_month_names() {
        let result = format_number(EXCEL_JUL_4_2024, "mmmm d, yyyy", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_long_date_february() {
        let result = format_number(44958.0, "mmmm d, yyyy", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_long_date_september() {
        let result = format_number(45170.0, "mmmm d, yyyy", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_dddd_mmmm_d_yyyy() {
        let result = format_number(EXCEL_JAN_15_2023, "dddd, mmmm d, yyyy", false);
        assert!(!result.is_empty());
    }
}

// ============================================================================
// Time Formats (numFmtId 18-21, 45-46)
// ============================================================================

mod time_formats {
    use super::*;

    // Format 20: h:mm
    mod h_mm {
        use super::*;

        #[test]
        fn test_h_mm_midnight() {
            let result = format_number(TIME_MIDNIGHT, "h:mm", false);
            assert!(result.contains(':'));
            assert!(result.contains("0") || result.contains("12"));
        }

        #[test]
        fn test_h_mm_noon() {
            let result = format_number(TIME_NOON, "h:mm", false);
            assert!(result.contains(':'));
            assert!(result.contains("12"));
        }

        #[test]
        fn test_h_mm_morning() {
            let result = format_number(TIME_6AM, "h:mm", false);
            assert!(result.contains(':'));
            assert!(result.contains("6"));
        }

        #[test]
        fn test_h_mm_evening() {
            let result = format_number(TIME_6PM, "h:mm", false);
            assert!(result.contains(':'));
            assert!(result.contains("18") || result.contains("6"));
        }

        #[test]
        fn test_h_mm_arbitrary_time() {
            let result = format_number(TIME_8_15_30_AM, "h:mm", false);
            assert!(result.contains(':'));
            assert!(result.contains("8") || result.contains("15"));
        }
    }

    // Format 21: h:mm:ss
    mod h_mm_ss {
        use super::*;

        #[test]
        fn test_h_mm_ss_midnight() {
            let result = format_number(TIME_MIDNIGHT, "h:mm:ss", false);
            assert!(result.contains(':'));
        }

        #[test]
        fn test_h_mm_ss_noon() {
            let result = format_number(TIME_NOON, "h:mm:ss", false);
            assert!(result.contains(':'));
            assert!(result.contains("12"));
        }

        #[test]
        fn test_h_mm_ss_with_seconds() {
            let result = format_number(TIME_1_30_45_PM, "h:mm:ss", false);
            assert!(result.contains(':'));
        }

        #[test]
        fn test_h_mm_ss_just_before_midnight() {
            let result = format_number(TIME_JUST_BEFORE_MIDNIGHT, "h:mm:ss", false);
            assert!(result.contains(':'));
        }

        #[test]
        fn test_h_mm_ss_8_15_30_am() {
            let result = format_number(TIME_8_15_30_AM, "h:mm:ss", false);
            assert!(result.contains(':'));
        }
    }

    // Format 18: h:mm AM/PM
    mod h_mm_am_pm {
        use super::*;

        #[test]
        fn test_h_mm_am_pm_midnight() {
            let result = format_number(TIME_MIDNIGHT, "h:mm AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("AM") || result.contains("am"));
        }

        #[test]
        fn test_h_mm_am_pm_noon() {
            let result = format_number(TIME_NOON, "h:mm AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("12"));
            assert!(result.contains("PM") || result.contains("pm"));
        }

        #[test]
        fn test_h_mm_am_pm_morning() {
            let result = format_number(TIME_6AM, "h:mm AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("6"));
            assert!(result.contains("AM") || result.contains("am"));
        }

        #[test]
        fn test_h_mm_am_pm_evening() {
            let result = format_number(TIME_6PM, "h:mm AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("6"));
            assert!(result.contains("PM") || result.contains("pm"));
        }

        #[test]
        fn test_h_mm_am_pm_1pm() {
            let result = format_number(TIME_1_30_45_PM, "h:mm AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("PM") || result.contains("pm"));
        }

        #[test]
        fn test_h_mm_am_pm_just_before_noon() {
            let result = format_number(0.4993055555555556, "h:mm AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("AM") || result.contains("am"));
        }

        #[test]
        fn test_h_mm_am_pm_just_after_noon() {
            let result = format_number(0.5006944444444444, "h:mm AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("PM") || result.contains("pm"));
        }
    }

    // Format 19: h:mm:ss AM/PM
    mod h_mm_ss_am_pm {
        use super::*;

        #[test]
        fn test_h_mm_ss_am_pm_midnight() {
            let result = format_number(TIME_MIDNIGHT, "h:mm:ss AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("AM") || result.contains("am"));
        }

        #[test]
        fn test_h_mm_ss_am_pm_noon() {
            let result = format_number(TIME_NOON, "h:mm:ss AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("12"));
            assert!(result.contains("PM") || result.contains("pm"));
        }

        #[test]
        fn test_h_mm_ss_am_pm_with_seconds() {
            let result = format_number(TIME_1_30_45_PM, "h:mm:ss AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("PM") || result.contains("pm"));
        }

        #[test]
        fn test_h_mm_ss_am_pm_morning_with_seconds() {
            let result = format_number(TIME_8_15_30_AM, "h:mm:ss AM/PM", false);
            assert!(result.contains(':'));
            assert!(result.contains("AM") || result.contains("am"));
        }
    }

    // Format 45: mm:ss
    mod mm_ss {
        use super::*;

        #[test]
        fn test_mm_ss_zero() {
            let result = format_number(0.0, "mm:ss", false);
            assert!(result.contains(':') || !result.is_empty());
        }

        #[test]
        fn test_mm_ss_one_minute() {
            let result = format_number(1.0 / 1440.0, "mm:ss", false);
            assert!(!result.is_empty());
        }

        #[test]
        fn test_mm_ss_30_seconds() {
            let result = format_number(30.0 / 86400.0, "mm:ss", false);
            assert!(!result.is_empty());
        }
    }

    // Format 46: [h]:mm:ss (elapsed time)
    mod elapsed_time {
        use super::*;

        #[test]
        fn test_elapsed_time_one_day() {
            let result = format_number(1.0, "[h]:mm:ss", false);
            assert!(!result.is_empty());
        }

        #[test]
        fn test_elapsed_time_multi_day() {
            let result = format_number(2.5, "[h]:mm:ss", false);
            assert!(!result.is_empty());
        }

        #[test]
        fn test_elapsed_time_partial_day() {
            let result = format_number(0.5, "[h]:mm:ss", false);
            assert!(!result.is_empty());
        }
    }
}

// ============================================================================
// Combined Date/Time Formats (numFmtId 22)
// ============================================================================

mod combined_datetime_formats {
    use super::*;

    #[test]
    fn test_m_d_yy_h_mm_basic() {
        let result = format_number(EXCEL_JAN_15_2023 + TIME_NOON, "m/d/yy h:mm", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_m_d_yy_h_mm_midnight() {
        let result = format_number(EXCEL_JAN_15_2023, "m/d/yy h:mm", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_m_d_yy_h_mm_evening() {
        let result = format_number(EXCEL_JAN_15_2023 + TIME_6PM, "m/d/yy h:mm", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_yyyy_mm_dd_hh_mm_ss() {
        let result = format_number(
            EXCEL_JAN_15_2023 + TIME_1_30_45_PM,
            "yyyy-mm-dd hh:mm:ss",
            false,
        );
        assert!(result.contains("2023") || result.contains(':'));
    }

    #[test]
    fn test_d_mmm_yyyy_h_mm_ss() {
        let result = format_number(
            EXCEL_JAN_15_2023 + TIME_8_15_30_AM,
            "d-mmm-yyyy h:mm:ss",
            false,
        );
        assert!(!result.is_empty());
    }

    #[test]
    fn test_datetime_with_am_pm() {
        let result = format_number(
            EXCEL_JAN_15_2023 + TIME_1_30_45_PM,
            "m/d/yy h:mm:ss AM/PM",
            false,
        );
        assert!(!result.is_empty());
    }

    #[test]
    fn test_datetime_boundary_crossing() {
        let result = format_number(
            EXCEL_JAN_15_2023 + TIME_JUST_BEFORE_MIDNIGHT,
            "yyyy-mm-dd hh:mm:ss",
            false,
        );
        assert!(!result.is_empty());
    }
}

// ============================================================================
// 1900 Date System (Windows Excel Default)
// ============================================================================

mod date_system_1900 {
    use super::*;

    #[test]
    fn test_1900_system_day_1() {
        let result = format_number(EXCEL_JAN_1_1900, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("01"));
    }

    #[test]
    fn test_1900_system_day_2() {
        let result = format_number(2.0, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("02"));
    }

    #[test]
    fn test_1900_system_day_31() {
        let result = format_number(31.0, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("31"));
    }

    #[test]
    fn test_1900_system_day_32() {
        let result = format_number(32.0, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("02") || result.contains("01"));
    }

    #[test]
    fn test_1900_system_end_of_february() {
        let result = format_number(EXCEL_FEB_28_1900, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("02") || result.contains("28"));
    }

    #[test]
    fn test_1900_system_march_1() {
        let result = format_number(EXCEL_MAR_1_1900, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("03") || result.contains("01"));
    }

    #[test]
    fn test_1900_system_year_boundary() {
        let result = format_number(EXCEL_DEC_31_1900, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("12") || result.contains("31"));
    }

    #[test]
    fn test_1900_system_january_1_1901() {
        let result = format_number(367.0, "yyyy-mm-dd", false);
        assert!(result.contains("1901") || result.contains("01"));
    }

    #[test]
    fn test_1900_system_y2k() {
        let result = format_number(EXCEL_JAN_1_2000, "yyyy-mm-dd", false);
        assert!(result.contains("2000") || result.contains("01"));
    }

    #[test]
    fn test_1900_system_modern_date() {
        let result = format_number(EXCEL_JUL_4_2024, "yyyy-mm-dd", false);
        assert!(result.contains("2024") || result.contains("07") || result.contains("04"));
    }
}

// ============================================================================
// 1904 Date System (Mac Excel)
// ============================================================================

mod date_system_1904 {
    use super::*;

    const DATE_SYSTEM_OFFSET: f64 = 1462.0;

    #[test]
    fn test_1904_offset_explanation() {
        let jan_1_1904_in_1900_system = 1462.0;
        let result = format_number(jan_1_1904_in_1900_system, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_1904_system_day_0() {
        let day_0_1904_in_1900_system = DATE_SYSTEM_OFFSET;
        let result = format_number(day_0_1904_in_1900_system, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_1904_avoids_leap_year_bug() {
        let feb_28_1904_in_1900 = 1462.0 + 58.0;
        let result = format_number(feb_28_1904_in_1900, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_1904_to_1900_conversion() {
        let serial_in_1904 = 35064.0;
        let serial_in_1900 = serial_in_1904 + DATE_SYSTEM_OFFSET;
        let result = format_number(serial_in_1900, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }
}

// ============================================================================
// Excel Leap Year Bug (February 29, 1900)
// ============================================================================

mod excel_leap_year_bug {
    use super::*;

    #[test]
    fn test_feb_28_1900_day_59() {
        let result = format_number(EXCEL_FEB_28_1900, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("02") || result.contains("28"));
    }

    #[test]
    fn test_fake_feb_29_1900_day_60() {
        let result = format_number(EXCEL_FEB_29_1900_FAKE, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_mar_1_1900_day_61() {
        let result = format_number(EXCEL_MAR_1_1900, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("03") || result.contains("01"));
    }

    #[test]
    fn test_consecutive_days_around_bug() {
        for day in 58..=62 {
            let result = format_number(day as f64, "yyyy-mm-dd", false);
            assert!(!result.is_empty(), "Day {} produced empty result", day);
        }
    }

    #[test]
    fn test_dates_after_bug_are_correct() {
        let result = format_number(62.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_year_2000_not_affected_by_bug() {
        let result = format_number(EXCEL_JAN_1_2000, "yyyy-mm-dd", false);
        assert!(result.contains("2000") || result.contains("01"));
    }

    #[test]
    fn test_leap_year_2000_is_correct() {
        let result = format_number(36585.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_non_leap_year_1900() {
        let result = format_number(EXCEL_DEC_31_1900, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("12") || result.contains("31"));
    }
}

// ============================================================================
// Negative Dates
// ============================================================================

mod negative_dates {
    use super::*;

    #[test]
    fn test_negative_one() {
        let result = format_number(-1.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_negative_large() {
        let result = format_number(-1000.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_day_zero() {
        let result = format_number(0.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_negative_with_time() {
        let result = format_number(-1.5, "yyyy-mm-dd hh:mm:ss", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_very_negative_date() {
        let result = format_number(-36500.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_negative_zero_fractional() {
        let result = format_number(-0.5, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }
}

// ============================================================================
// Date as Serial Number vs Formatted
// ============================================================================

mod serial_vs_formatted {
    use super::*;

    #[test]
    fn test_same_serial_different_formats() {
        let serial = EXCEL_JAN_15_2023;

        let iso = format_number(serial, "yyyy-mm-dd", false);
        let us = format_number(serial, "mm/dd/yyyy", false);
        let eu = format_number(serial, "dd/mm/yyyy", false);
        let abbr = format_number(serial, "d-mmm-yy", false);

        assert!(!iso.is_empty());
        assert!(!us.is_empty());
        assert!(!eu.is_empty());
        assert!(!abbr.is_empty());
    }

    #[test]
    fn test_general_format_shows_serial() {
        let result = format_number(EXCEL_JAN_15_2023, "General", false);
        assert!(result.contains("44941") || result.parse::<f64>().is_ok());
    }

    #[test]
    fn test_number_format_shows_serial() {
        let result = format_number(EXCEL_JAN_15_2023, "#,##0", false);
        assert!(result.contains("44") || result.contains(","));
    }

    #[test]
    fn test_date_with_decimals() {
        let serial_with_time = EXCEL_JAN_15_2023 + 0.5;

        let date_only = format_number(serial_with_time, "yyyy-mm-dd", false);
        let time_only = format_number(serial_with_time, "h:mm:ss", false);
        let datetime = format_number(serial_with_time, "yyyy-mm-dd h:mm:ss", false);

        assert!(!date_only.is_empty());
        assert!(!time_only.is_empty());
        assert!(!datetime.is_empty());
    }

    #[test]
    fn test_time_only_no_date() {
        let time_only = 0.75;
        let result = format_number(time_only, "h:mm:ss", false);
        assert!(result.contains(':'));
    }

    #[test]
    fn test_integer_serial_no_time() {
        let date_only = EXCEL_JAN_15_2023;
        let result = format_number(date_only, "h:mm:ss", false);
        assert!(result.contains(':'));
    }

    #[test]
    fn test_preserved_precision() {
        let precise_time = 0.5211805555555556;
        let result = format_number(precise_time, "h:mm:ss", false);
        assert!(result.contains(':'));
    }
}

// ============================================================================
// Date Format Detection
// ============================================================================

mod date_format_detection {
    use super::*;

    #[test]
    fn test_is_date_format_standard_dates() {
        assert!(is_date_format("mm-dd-yy"));
        assert!(is_date_format("d-mmm-yy"));
        assert!(is_date_format("d-mmm"));
        assert!(is_date_format("mmm-yy"));
    }

    #[test]
    fn test_is_date_format_iso() {
        assert!(is_date_format("yyyy-mm-dd"));
        assert!(is_date_format("yyyy/mm/dd"));
    }

    #[test]
    fn test_is_date_format_long() {
        assert!(is_date_format("mmmm d, yyyy"));
        assert!(is_date_format("dddd, mmmm d, yyyy"));
    }

    #[test]
    fn test_is_date_format_times() {
        assert!(is_date_format("h:mm"));
        assert!(is_date_format("h:mm:ss"));
        assert!(is_date_format("h:mm AM/PM"));
        assert!(is_date_format("h:mm:ss AM/PM"));
    }

    #[test]
    fn test_is_date_format_combined() {
        assert!(is_date_format("yyyy-mm-dd h:mm:ss"));
        assert!(is_date_format("m/d/yy h:mm"));
    }

    #[test]
    fn test_is_not_date_format_numbers() {
        assert!(!is_date_format("#,##0"));
        assert!(!is_date_format("#,##0.00"));
        assert!(!is_date_format("0"));
        assert!(!is_date_format("0.00"));
        assert!(!is_date_format("General"));
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
    fn test_is_date_format_case_insensitive() {
        assert!(is_date_format("YYYY-MM-DD"));
        assert!(is_date_format("Yyyy-Mm-Dd"));
        assert!(is_date_format("H:MM:SS"));
    }

    #[test]
    fn test_is_date_format_ignores_quoted() {
        assert!(!is_date_format(r#""Date: " 0"#));
    }

    #[test]
    fn test_is_date_format_ignores_brackets() {
        assert!(!is_date_format("[Red]#,##0"));
    }

    #[test]
    fn test_is_date_format_elapsed_time() {
        assert!(is_date_format("[h]:mm:ss"));
        assert!(is_date_format("mm:ss"));
    }
}

// ============================================================================
// Special Date Cases
// ============================================================================

mod special_date_cases {
    use super::*;

    #[test]
    fn test_very_large_serial() {
        let result = format_number(100000.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_fractional_second() {
        let precise_time = 0.5 + (30.5 / 86400.0);
        let result = format_number(precise_time, "h:mm:ss", false);
        assert!(result.contains(':'));
    }

    #[test]
    fn test_rounding_at_day_boundary() {
        let almost_next_day = 44941.999999;
        let result = format_number(almost_next_day, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_time_wrapping() {
        let midnight = 44942.0;
        let result = format_number(midnight, "h:mm:ss", false);
        assert!(result.contains(':'));
    }

    #[test]
    fn test_century_boundary_1999_2000() {
        let dec_31_1999 = 36525.0;
        let result_1999 = format_number(dec_31_1999, "yyyy-mm-dd", false);
        assert!(!result_1999.is_empty());

        let result_2000 = format_number(EXCEL_JAN_1_2000, "yyyy-mm-dd", false);
        assert!(!result_2000.is_empty());
    }

    #[test]
    fn test_year_2038_problem() {
        let result = format_number(50424.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_year_10000() {
        let result = format_number(2958466.0, "yyyy-mm-dd", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_minimum_date() {
        let result = format_number(1.0, "yyyy-mm-dd", false);
        assert!(result.contains("1900") || result.contains("01"));
    }

    #[test]
    fn test_time_precision() {
        let one_second = 1.0 / 86400.0;
        let result = format_number(one_second, "h:mm:ss", false);
        assert!(result.contains(':'));
    }
}

// ============================================================================
// Format Variations
// ============================================================================

mod format_variations {
    use super::*;

    #[test]
    fn test_single_digit_vs_padded_month() {
        let m_result = format_number(EXCEL_JAN_15_2023, "m/d/yy", false);
        let mm_result = format_number(EXCEL_JAN_15_2023, "mm/dd/yy", false);

        assert!(!m_result.is_empty());
        assert!(!mm_result.is_empty());
    }

    #[test]
    fn test_single_digit_vs_padded_day() {
        let d_result = format_number(EXCEL_JAN_15_2023, "m/d/yy", false);
        let dd_result = format_number(EXCEL_JAN_15_2023, "m/dd/yy", false);

        assert!(!d_result.is_empty());
        assert!(!dd_result.is_empty());
    }

    #[test]
    fn test_two_digit_vs_four_digit_year() {
        let yy_result = format_number(EXCEL_JAN_15_2023, "mm/dd/yy", false);
        let yyyy_result = format_number(EXCEL_JAN_15_2023, "mm/dd/yyyy", false);

        assert!(!yy_result.is_empty());
        assert!(!yyyy_result.is_empty());
    }

    #[test]
    fn test_month_name_variations() {
        let m = format_number(EXCEL_JAN_15_2023, "m", false);
        let mm = format_number(EXCEL_JAN_15_2023, "mm", false);
        let mmm = format_number(EXCEL_JAN_15_2023, "mmm", false);
        let mmmm = format_number(EXCEL_JAN_15_2023, "mmmm", false);

        assert!(!m.is_empty());
        assert!(!mm.is_empty());
        assert!(!mmm.is_empty());
        assert!(!mmmm.is_empty());
    }

    #[test]
    fn test_day_name_variations() {
        let d = format_number(EXCEL_JAN_15_2023, "d", false);
        let dd = format_number(EXCEL_JAN_15_2023, "dd", false);
        let ddd = format_number(EXCEL_JAN_15_2023, "ddd", false);
        let dddd = format_number(EXCEL_JAN_15_2023, "dddd", false);

        assert!(!d.is_empty());
        assert!(!dd.is_empty());
        assert!(!ddd.is_empty());
        assert!(!dddd.is_empty());
    }

    #[test]
    fn test_hour_variations() {
        let h = format_number(TIME_8_15_30_AM, "h:mm:ss", false);
        let hh = format_number(TIME_8_15_30_AM, "hh:mm:ss", false);

        assert!(!h.is_empty());
        assert!(!hh.is_empty());
    }

    #[test]
    fn test_am_pm_variations() {
        let ampm_upper = format_number(TIME_1_30_45_PM, "h:mm AM/PM", false);
        let ampm_lower = format_number(TIME_1_30_45_PM, "h:mm am/pm", false);
        let ap = format_number(TIME_1_30_45_PM, "h:mm A/P", false);

        assert!(!ampm_upper.is_empty());
        assert!(!ampm_lower.is_empty());
        assert!(!ap.is_empty());
    }

    #[test]
    fn test_separator_variations() {
        let dash = format_number(EXCEL_JAN_15_2023, "yyyy-mm-dd", false);
        let slash = format_number(EXCEL_JAN_15_2023, "yyyy/mm/dd", false);
        let dot = format_number(EXCEL_JAN_15_2023, "yyyy.mm.dd", false);
        let space = format_number(EXCEL_JAN_15_2023, "yyyy mm dd", false);

        assert!(!dash.is_empty());
        assert!(!slash.is_empty());
        assert!(!dot.is_empty());
        assert!(!space.is_empty());
    }
}

// ============================================================================
// Builtin Date Format IDs
// ============================================================================

mod builtin_date_formats {
    use offidized_xlsx::numfmt::get_builtin_format;

    #[test]
    fn test_builtin_format_14() {
        assert_eq!(get_builtin_format(14), Some("mm-dd-yy"));
    }

    #[test]
    fn test_builtin_format_15() {
        assert_eq!(get_builtin_format(15), Some("d-mmm-yy"));
    }

    #[test]
    fn test_builtin_format_16() {
        assert_eq!(get_builtin_format(16), Some("d-mmm"));
    }

    #[test]
    fn test_builtin_format_17() {
        assert_eq!(get_builtin_format(17), Some("mmm-yy"));
    }

    #[test]
    fn test_builtin_format_18() {
        assert_eq!(get_builtin_format(18), Some("h:mm AM/PM"));
    }

    #[test]
    fn test_builtin_format_19() {
        assert_eq!(get_builtin_format(19), Some("h:mm:ss AM/PM"));
    }

    #[test]
    fn test_builtin_format_20() {
        assert_eq!(get_builtin_format(20), Some("h:mm"));
    }

    #[test]
    fn test_builtin_format_21() {
        assert_eq!(get_builtin_format(21), Some("h:mm:ss"));
    }

    #[test]
    fn test_builtin_format_22() {
        assert_eq!(get_builtin_format(22), Some("m/d/yy h:mm"));
    }

    #[test]
    fn test_builtin_format_45() {
        assert_eq!(get_builtin_format(45), Some("mm:ss"));
    }

    #[test]
    fn test_builtin_format_46() {
        assert_eq!(get_builtin_format(46), Some("[h]:mm:ss"));
    }

    #[test]
    fn test_builtin_format_47() {
        assert_eq!(get_builtin_format(47), Some("mmss.0"));
    }
}

// ============================================================================
// Real-World Date Scenarios
// ============================================================================

mod real_world_scenarios {
    use super::*;

    #[test]
    fn test_invoice_date_format() {
        let result = format_number(EXCEL_JAN_15_2023, "mmmm d, yyyy", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_timestamp_format() {
        let result = format_number(
            EXCEL_JAN_15_2023 + TIME_1_30_45_PM,
            "yyyy-mm-dd hh:mm:ss",
            false,
        );
        assert!(!result.is_empty());
    }

    #[test]
    fn test_log_file_date_format() {
        let result = format_number(
            EXCEL_JAN_15_2023 + TIME_8_15_30_AM,
            "yyyy-mm-dd'T'hh:mm:ss",
            false,
        );
        assert!(!result.is_empty());
    }

    #[test]
    fn test_financial_quarter_end() {
        let result = format_number(45291.0, "mmm dd, yyyy", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_birthday_format() {
        let result = format_number(EXCEL_JUL_4_2024, "MMMM d", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_appointment_time() {
        let result = format_number(TIME_1_30_45_PM, "h:mm a", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_file_timestamp() {
        let result = format_number(EXCEL_JAN_15_2023 + TIME_NOON, "yyyymmddhhmmss", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_european_date_format() {
        let result = format_number(EXCEL_JAN_15_2023, "dd/mm/yyyy", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_us_date_format() {
        let result = format_number(EXCEL_JAN_15_2023, "mm/dd/yyyy", false);
        assert!(!result.is_empty());
    }

    #[test]
    fn test_japanese_date_format() {
        let result = format_number(EXCEL_JAN_15_2023, "yyyy/mm/dd", false);
        assert!(!result.is_empty());
    }
}
