//! Date and time functions.
//!
//! Excel uses a serial date number system where day 1 = January 1, 1900.
//! Excel incorrectly treats 1900 as a leap year: serial 60 = Feb 29, 1900
//! (a date that never existed). We preserve this bug for compatibility.

use crate::context::EvalContext;
use crate::value::{CalcValue, ScalarValue, XlError};

use super::{
    iter_range_values, require_number, require_scalar, require_text, FunctionArg, FunctionDef,
    FunctionRegistry, ParamKind,
};

/// Serial date of 9999-12-31. Date is generally considered invalid if above this or below 0.
const YEAR_10K: i64 = 2958465;

/// `TODAY()`
///
/// Returns the current date as a serial number. Currently returns 0.0 as a
/// placeholder (system date access will be routed through the provider in a
/// future phase).
pub fn fn_today(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::number(0.0)
}

/// `NOW()`
///
/// Returns the current date and time as a serial number with fractional time.
/// Currently returns 0.0 as a placeholder.
pub fn fn_now(_args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    CalcValue::number(0.0)
}

/// `DATE(year, month, day)`
///
/// Constructs an Excel serial date number from year, month, and day components.
///
/// - Year 0-1899 has 1900 added.
/// - Month and day overflow wraps (e.g. month 13 = Jan of next year).
/// - Negative day counts backward.
pub fn fn_date(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let year = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n.floor() as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let month = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n.floor() as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let day = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => n.floor() as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    // Year adjustment: 0-1899 has 1900 added
    let adjusted_year = if year < 1900 { year + 1900 } else { year };

    if adjusted_year > 10000 {
        return CalcValue::error(XlError::Num);
    }

    match date_to_serial(adjusted_year, month, day) {
        Some(serial) if serial > 0 && serial < YEAR_10K => CalcValue::number(serial as f64),
        _ => CalcValue::error(XlError::Num),
    }
}

/// `YEAR(serial_number)`
///
/// Extracts the year from an Excel serial date number.
pub fn fn_year(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if !(0..=YEAR_10K).contains(&serial) {
        return CalcValue::error(XlError::Num);
    }
    let (y, _m, _d) = serial_to_date(serial);
    CalcValue::number(y as f64)
}

/// `MONTH(serial_number)`
///
/// Extracts the month (1-12) from an Excel serial date number.
pub fn fn_month(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if !(0..=YEAR_10K).contains(&serial) {
        return CalcValue::error(XlError::Num);
    }
    let (_y, m, _d) = serial_to_date(serial);
    CalcValue::number(m as f64)
}

/// `DAY(serial_number)`
///
/// Extracts the day (1-31) from an Excel serial date number.
pub fn fn_day(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if !(0..=YEAR_10K).contains(&serial) {
        return CalcValue::error(XlError::Num);
    }
    let (_y, _m, d) = serial_to_date(serial);
    CalcValue::number(d as f64)
}

/// `HOUR(serial_time)`
///
/// Extracts the hour (0-23) from an Excel serial time number.
pub fn fn_hour(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if serial < 0.0 || serial as i64 >= YEAR_10K {
        return CalcValue::error(XlError::Num);
    }
    let time_part = serial - serial.floor();
    let total_seconds = (time_part * 86400.0).round() as i64;
    let hour = (total_seconds / 3600) % 24;
    CalcValue::number(hour as f64)
}

/// `MINUTE(serial_time)`
///
/// Extracts the minute (0-59) from an Excel serial time number.
pub fn fn_minute(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if serial < 0.0 || serial as i64 >= YEAR_10K {
        return CalcValue::error(XlError::Num);
    }
    let time_part = serial - serial.floor();
    let total_seconds = (time_part * 86400.0).round() as i64;
    let minute = (total_seconds / 60) % 60;
    CalcValue::number(minute as f64)
}

/// `SECOND(serial_time)`
///
/// Extracts the second (0-59) from an Excel serial time number.
pub fn fn_second(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    if serial < 0.0 || serial as i64 >= YEAR_10K {
        return CalcValue::error(XlError::Num);
    }
    let time_part = serial - serial.floor();
    let total_seconds = (time_part * 86400.0).round() as i64;
    let second = total_seconds % 60;
    CalcValue::number(second as f64)
}

/// `TIME(hour, minute, second)`
///
/// Constructs a time serial number from hour, minute, and second components.
pub fn fn_time(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let hour = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n.floor() as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let minute = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n.floor() as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let second = match args.get(2) {
        Some(a) => match require_number(a) {
            Ok(n) => n.floor() as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if !(0..=32767).contains(&hour)
        || !(0..=32767).contains(&minute)
        || !(0..=32767).contains(&second)
    {
        return CalcValue::error(XlError::Num);
    }

    let total_seconds = hour * 3600 + minute * 60 + second;
    let serial = (total_seconds as f64) / 86400.0;
    CalcValue::number(serial % 1.0)
}

/// `DAYS(end_date, start_date)`
///
/// Returns the number of days between two dates.
pub fn fn_days(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let end_date = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let start_date = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if !(0..=YEAR_10K).contains(&start_date) || !(0..=YEAR_10K).contains(&end_date) {
        return CalcValue::error(XlError::Num);
    }

    CalcValue::number((end_date - start_date) as f64)
}

/// `EDATE(start_date, months)`
///
/// Returns the serial number of the date that is the indicated number of months
/// before or after the start date.
pub fn fn_edate(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let start_serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let months = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if !(0..=YEAR_10K).contains(&start_serial) {
        return CalcValue::error(XlError::Num);
    }
    if !(-9999 * 12..=9999 * 12).contains(&months) {
        return CalcValue::error(XlError::Num);
    }

    let (y, m, d) = serial_to_date(start_serial);
    let new_month = m + months;
    let year_adjust = if new_month > 0 {
        (new_month - 1) / 12
    } else {
        (new_month - 12) / 12
    };
    let new_year = y + year_adjust;
    let final_month = new_month - year_adjust * 12;
    let final_day = d.min(days_in_month(new_year, final_month));

    match date_to_serial(new_year, final_month, final_day) {
        Some(serial) if serial > 0 && serial < YEAR_10K => CalcValue::number(serial as f64),
        _ => CalcValue::error(XlError::Num),
    }
}

/// `EOMONTH(start_date, months)`
///
/// Returns the serial number of the last day of the month before or after
/// a specified number of months.
pub fn fn_eomonth(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let start_serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let months = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if !(0..=YEAR_10K).contains(&start_serial) {
        return CalcValue::error(XlError::Num);
    }
    if !(-9999 * 12..=9999 * 12).contains(&months) {
        return CalcValue::error(XlError::Num);
    }

    let (y, m, _d) = serial_to_date(start_serial);
    let new_month = m + months;
    let year_adjust = if new_month > 0 {
        (new_month - 1) / 12
    } else {
        (new_month - 12) / 12
    };
    let new_year = y + year_adjust;
    let final_month = new_month - year_adjust * 12;
    let final_day = days_in_month(new_year, final_month);

    match date_to_serial(new_year, final_month, final_day) {
        Some(serial) if serial > 0 && serial < YEAR_10K => CalcValue::number(serial as f64),
        _ => CalcValue::error(XlError::Num),
    }
}

/// `WEEKDAY(serial_number, [return_type])`
///
/// Converts a serial number to a day of the week.
pub fn fn_weekday(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let return_type = if let Some(arg) = args.get(1) {
        match require_number(arg) {
            Ok(n) => n as i32,
            Err(e) => return e,
        }
    } else {
        1
    };

    if !(0..=YEAR_10K).contains(&serial) {
        return CalcValue::error(XlError::Num);
    }

    let week_start_offset = match return_type {
        1 => 0,  // Sun
        2 => 6,  // Mon
        3 => 6,  // Mon
        11 => 6, // Mon
        12 => 5, // Tue
        13 => 4, // Wed
        14 => 3, // Thu
        15 => 2, // Fri
        16 => 1, // Sat
        17 => 0, // Sun
        _ => return CalcValue::error(XlError::Num),
    };

    let number_offset = if return_type == 3 { 0 } else { 1 };
    let weekday = (serial + 6 + week_start_offset) % 7 + number_offset;
    CalcValue::number(weekday as f64)
}

/// `WEEKNUM(serial_number, [return_type])`
///
/// Converts a serial number to a number representing where the week falls
/// numerically within a year.
pub fn fn_weeknum(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let return_type = if let Some(arg) = args.get(1) {
        match require_number(arg) {
            Ok(n) => n as i32,
            Err(e) => return e,
        }
    } else {
        1
    };

    if !(0..=YEAR_10K).contains(&serial) {
        return CalcValue::error(XlError::Num);
    }

    // Special case: ISO week (21)
    if return_type == 21 {
        return fn_isoweeknum(args, _ctx);
    }

    let first_day_of_week = match return_type {
        1 => 0,  // Sunday
        2 => 1,  // Monday
        11 => 1, // Monday
        12 => 2, // Tuesday
        13 => 3, // Wednesday
        14 => 4, // Thursday
        15 => 5, // Friday
        16 => 6, // Saturday
        17 => 0, // Sunday
        _ => return CalcValue::error(XlError::Num),
    };

    if serial == 0 && first_day_of_week == 0 {
        return CalcValue::number(0.0);
    }

    let (year, _, _) = serial_to_date(serial);
    let start_of_year_serial = date_to_serial(year, 1, 1).unwrap_or(1);
    let start_of_year_day_of_week = (start_of_year_serial + 6) % 7;
    let mut start_of_week_adjust = first_day_of_week - start_of_year_day_of_week;

    if start_of_week_adjust > 0 {
        start_of_week_adjust -= 7;
    }

    let first_week_start_date = start_of_year_serial + start_of_week_adjust;
    let week_num = (serial - first_week_start_date) / 7;
    CalcValue::number((week_num + 1) as f64)
}

/// `ISOWEEKNUM(serial_number)`
///
/// Returns the ISO week number of the year for a given date.
pub fn fn_isoweeknum(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if !(0..=YEAR_10K).contains(&serial) {
        return CalcValue::error(XlError::Num);
    }

    let (year, _, _) = serial_to_date(serial);
    let day_of_year = serial - date_to_serial(year, 1, 1).unwrap_or(0) + 1;
    let day_of_week = ((serial + 6) % 7 + 1) as i64; // Monday = 1, Sunday = 7
    let week = (10 + day_of_year - day_of_week) / 7;

    if week < 1 {
        return CalcValue::number(weeks_in_iso_year(year - 1) as f64);
    }

    if week > weeks_in_iso_year(year) {
        return CalcValue::number(1.0);
    }

    CalcValue::number(week as f64)
}

fn weeks_in_iso_year(year: i64) -> i64 {
    let dec31_day_of_week = (year + year / 4 - year / 100 + year / 400) % 7;
    if dec31_day_of_week == 4 {
        return 53;
    }
    let dec31_prev_year = ((year - 1) + (year - 1) / 4 - (year - 1) / 100 + (year - 1) / 400) % 7;
    if dec31_prev_year == 3 {
        return 53;
    }
    52
}

/// `DATEVALUE(date_text)`
///
/// Converts a date in the form of text to a serial number.
pub fn fn_datevalue(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let _text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    // TODO: Implement date parsing logic
    CalcValue::error(XlError::Value)
}

/// `TIMEVALUE(time_text)`
///
/// Converts a time in the form of text to a serial number.
pub fn fn_timevalue(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let _text = match args.first() {
        Some(a) => match require_text(a) {
            Ok(s) => s,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    // TODO: Implement time parsing logic
    CalcValue::error(XlError::Value)
}

/// `DAYS360(start_date, end_date, [method])`
///
/// Calculates the number of days between two dates based on a 360-day year.
pub fn fn_days360(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let start_serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let end_serial = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let is_european = if let Some(arg) = args.get(2) {
        match require_scalar(arg) {
            ScalarValue::Bool(b) => *b,
            ScalarValue::Number(n) => *n != 0.0,
            ScalarValue::Blank => false,
            ScalarValue::Error(e) => return CalcValue::error(*e),
            _ => return CalcValue::error(XlError::Value),
        }
    } else {
        false
    };

    if !(0..=YEAR_10K).contains(&start_serial) || !(0..=YEAR_10K).contains(&end_serial) {
        return CalcValue::error(XlError::Num);
    }

    let (start_y, start_m, mut start_d) = serial_to_date(start_serial);
    let (end_y, end_m, mut end_d) = serial_to_date(end_serial);

    if is_european {
        if start_d == 31 {
            start_d = 30;
        }
        if end_d == 31 {
            end_d = 30;
        }
    } else {
        let start_is_last_day = start_d == days_in_month(start_y, start_m);
        if start_is_last_day {
            start_d = 30;
        }
        if end_d == 31 && start_d == 30 {
            end_d = 30;
        }
    }

    let days = 360 * (end_y - start_y) + 30 * (end_m - start_m) + (end_d - start_d);
    CalcValue::number(days as f64)
}

/// `DATEDIF(start_date, end_date, unit)`
///
/// Calculates the number of days, months, or years between two dates.
pub fn fn_datedif(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let start_serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let end_serial = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let unit = match args.get(2) {
        Some(a) => match require_text(a) {
            Ok(s) => s.to_uppercase(),
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if !(0..=YEAR_10K).contains(&start_serial) || !(0..=YEAR_10K).contains(&end_serial) {
        return CalcValue::error(XlError::Num);
    }

    if start_serial > end_serial {
        return CalcValue::error(XlError::Num);
    }

    let (start_y, start_m, start_d) = serial_to_date(start_serial);
    let (end_y, end_m, end_d) = serial_to_date(end_serial);

    match unit.as_str() {
        "Y" => {
            let is_last_year_complete = end_m > start_m || (end_m == start_m && end_d >= start_d);
            let years = end_y - start_y - if is_last_year_complete { 0 } else { 1 };
            CalcValue::number(years as f64)
        }
        "M" => {
            let is_last_month_complete = end_d >= start_d;
            let months = (end_y - start_y) * 12 + end_m
                - start_m
                - if is_last_month_complete { 0 } else { 1 };
            CalcValue::number(months as f64)
        }
        "D" => CalcValue::number((end_serial - start_serial) as f64),
        "MD" => {
            if end_d >= start_d {
                CalcValue::number((end_d - start_d) as f64)
            } else {
                let adj_month = if end_m > 1 { end_m - 1 } else { 12 };
                let adj_year = if end_m > 1 { end_y } else { end_y - 1 };
                let adj_serial = date_to_serial(adj_year, adj_month, start_d).unwrap_or(0);
                CalcValue::number((end_serial - adj_serial) as f64)
            }
        }
        "YM" => {
            let is_last_month_complete = end_d >= start_d;
            let months = (end_m + 12 - start_m - if is_last_month_complete { 0 } else { 1 }) % 12;
            CalcValue::number(months as f64)
        }
        "YD" => {
            let end_follows_start = end_m > start_m || (end_m == start_m && end_d >= start_d);
            let new_end_year = start_y + if end_follows_start { 0 } else { 1 };
            let new_end_serial = date_to_serial(new_end_year, end_m, end_d).unwrap_or(0);
            let mut days_diff = new_end_serial - start_serial;
            if start_serial <= 60 && end_y > 1900 && end_m == 3 && end_d < start_d {
                days_diff -= 1;
            }
            CalcValue::number(days_diff as f64)
        }
        _ => CalcValue::error(XlError::Num),
    }
}

/// `NETWORKDAYS(start_date, end_date, [holidays])`
///
/// Returns the number of whole workdays between two dates.
pub fn fn_networkdays(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let start_serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let end_serial = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if !(0..=YEAR_10K).contains(&start_serial) || !(0..=YEAR_10K).contains(&end_serial) {
        return CalcValue::error(XlError::Num);
    }

    let mut holidays = Vec::new();
    if let Some(arg) = args.get(2) {
        match arg {
            FunctionArg::Value(v) => {
                let scalar = v.as_scalar();
                if let ScalarValue::Error(e) = scalar {
                    return CalcValue::error(*e);
                }
                if let ScalarValue::Number(n) = scalar.to_number() {
                    let h = n as i64;
                    if (0..=YEAR_10K).contains(&h) {
                        holidays.push(h);
                    }
                }
            }
            FunctionArg::Range { range, ctx: rctx } => {
                for val in iter_range_values(range, rctx) {
                    if let ScalarValue::Error(e) = val {
                        return CalcValue::error(e);
                    }
                    if let ScalarValue::Number(n) = val.to_number() {
                        let h = n as i64;
                        if (0..=YEAR_10K).contains(&h) {
                            holidays.push(h);
                        }
                    }
                }
            }
        }
    }

    holidays.sort_unstable();
    holidays.dedup();

    let result = business_days_until(start_serial, end_serial, &holidays);
    CalcValue::number(result as f64)
}

/// `WORKDAY(start_date, days, [holidays])`
///
/// Returns the serial number of the date before or after a specified number of workdays.
pub fn fn_workday(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let start_serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let day_offset = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    if !(0..=YEAR_10K).contains(&start_serial) {
        return CalcValue::error(XlError::Num);
    }

    if day_offset == 0 {
        return CalcValue::number(start_serial as f64);
    }

    let mut holidays = Vec::new();
    if let Some(arg) = args.get(2) {
        match arg {
            FunctionArg::Value(v) => {
                let scalar = v.as_scalar();
                if let ScalarValue::Error(e) = scalar {
                    return CalcValue::error(*e);
                }
                if let ScalarValue::Number(n) = scalar.to_number() {
                    let h = n as i64;
                    if (0..=YEAR_10K).contains(&h) {
                        holidays.push(h);
                    }
                }
            }
            FunctionArg::Range { range, ctx: rctx } => {
                for val in iter_range_values(range, rctx) {
                    if let ScalarValue::Error(e) = val {
                        return CalcValue::error(e);
                    }
                    if let ScalarValue::Number(n) = val.to_number() {
                        let h = n as i64;
                        if (0..=YEAR_10K).contains(&h) && !is_weekend(h) {
                            holidays.push(h);
                        }
                    }
                }
            }
        }
    }

    let direction = if day_offset > 0 { 1 } else { -1 };
    holidays.sort_unstable();
    if direction < 0 {
        holidays.reverse();
    }
    holidays.dedup();

    let holidays_filtered: Vec<i64> = holidays
        .into_iter()
        .filter(|&h| {
            if direction > 0 {
                h > start_serial
            } else {
                h < start_serial
            }
        })
        .collect();

    let mut last_date = start_serial;
    let mut workdays_so_far = 0;

    for &holiday in &holidays_filtered {
        let segment_workdays = if last_date + direction != holiday {
            business_days_until(last_date + direction, holiday, &[])
        } else {
            direction
        };

        if (direction > 0 && workdays_so_far + segment_workdays > day_offset)
            || (direction < 0 && workdays_so_far + segment_workdays < day_offset)
        {
            break;
        }

        workdays_so_far += segment_workdays - direction;
        last_date = holiday;
    }

    let remaining_workdays = day_offset - workdays_so_far;
    let week_count = remaining_workdays / 5;
    let mut remaining = remaining_workdays % 5;

    let mut week_adjust = week_count;
    if remaining == 0 {
        week_adjust -= direction;
        remaining += direction * 5;
    }

    let mut workday = last_date + week_adjust * 7;
    while remaining != 0 {
        loop {
            workday += direction;
            if !is_weekend(workday) {
                break;
            }
        }
        remaining -= direction;
    }

    CalcValue::number(workday as f64)
}

/// `YEARFRAC(start_date, end_date, [basis])`
///
/// Returns the year fraction representing the number of whole days between start_date and end_date.
pub fn fn_yearfrac(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let start_serial = match args.first() {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };
    let end_serial = match args.get(1) {
        Some(a) => match require_number(a) {
            Ok(n) => n as i64,
            Err(e) => return e,
        },
        None => return CalcValue::error(XlError::Value),
    };

    let basis = if let Some(arg) = args.get(2) {
        match require_number(arg) {
            Ok(n) => n as i32,
            Err(e) => return e,
        }
    } else {
        0
    };

    if !(0..=YEAR_10K).contains(&start_serial) || !(0..=YEAR_10K).contains(&end_serial) {
        return CalcValue::error(XlError::Num);
    }

    if !(0..5).contains(&basis) {
        return CalcValue::error(XlError::Num);
    }

    let year_frac = match basis {
        0 => {
            // US 30/360
            let (start_y, start_m, mut start_d) = serial_to_date(start_serial);
            let (end_y, end_m, mut end_d) = serial_to_date(end_serial);
            let start_is_last_day = start_d == days_in_month(start_y, start_m);
            if start_is_last_day {
                start_d = 30;
            }
            if end_d == 31 && start_d == 30 {
                end_d = 30;
            }
            let days = 360 * (end_y - start_y) + 30 * (end_m - start_m) + (end_d - start_d);
            days as f64 / 360.0
        }
        1 => {
            // Actual/Actual
            let (start_y, _, _) = serial_to_date(start_serial);
            let (end_y, _, _) = serial_to_date(end_serial);
            let mut total_days = 0;
            for year in start_y..=end_y {
                total_days += if is_real_leap_year(year) { 366 } else { 365 };
            }
            let year_avg = total_days as f64 / ((end_y - start_y + 1) as f64);
            (end_serial - start_serial) as f64 / year_avg
        }
        2 => {
            // Actual/360
            (end_serial - start_serial) as f64 / 360.0
        }
        3 => {
            // Actual/365
            (end_serial - start_serial) as f64 / 365.0
        }
        _ => {
            // EU 30/360
            let (start_y, start_m, mut start_d) = serial_to_date(start_serial);
            let (end_y, end_m, mut end_d) = serial_to_date(end_serial);
            if start_d == 31 {
                start_d = 30;
            }
            if end_d == 31 {
                end_d = 30;
            }
            let days = 360 * (end_y - start_y) + 30 * (end_m - start_m) + (end_d - start_d);
            days as f64 / 360.0
        }
    };

    CalcValue::number(year_frac.abs())
}

// ---------------------------------------------------------------------------
// Helper functions
// ---------------------------------------------------------------------------

/// Returns true if a serial date falls on a weekend (Saturday or Sunday).
fn is_weekend(serial: i64) -> bool {
    let weekday = (serial + 6) % 7 + 1; // 1=Sun, 7=Sat
    weekday == 1 || weekday == 7
}

/// Calculates the number of business days between two dates, excluding weekends and holidays.
fn business_days_until(first_day: i64, last_day: i64, holidays: &[i64]) -> i64 {
    if first_day > last_day {
        return -business_days_until(last_day, first_day, holidays);
    }

    let work_days = last_day - first_day + 1;
    let full_week_count = work_days / 7;
    let remaining_days = work_days % 7;

    let mut total = work_days;

    // Count weekends in remaining days
    for day in (last_day - remaining_days + 1)..=last_day {
        if is_weekend(day) {
            total -= 1;
        }
    }

    // Subtract weekends in full weeks
    total -= full_week_count * 2;

    // Subtract holidays
    for &holiday in holidays {
        if first_day <= holiday && holiday <= last_day && !is_weekend(holiday) {
            total -= 1;
        }
    }

    total
}

// ---------------------------------------------------------------------------
// Serial date <-> calendar date conversion
// ---------------------------------------------------------------------------
//
// Excel date system (1900-based):
//   serial 1  = January 1, 1900
//   serial 59 = February 28, 1900
//   serial 60 = February 29, 1900  (phantom leap day — 1900 was NOT a leap year)
//   serial 61 = March 1, 1900
//
// Strategy:
//   - Handle serials 1-60 (Jan 1 - Feb 29 1900) as a special case
//   - For serials > 60, subtract 1 to skip the phantom leap day, then
//     convert from a base of January 1, 1900 = day 0

/// Returns true if a year is a real leap year (not Excel's fake 1900 leap year).
fn is_real_leap_year(year: i64) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}

/// Days in a given real month (1-12) for a given year.
fn days_in_month(year: i64, month: i64) -> i64 {
    match month {
        1 => 31,
        2 => {
            if is_real_leap_year(year) {
                29
            } else {
                28
            }
        }
        3 => 31,
        4 => 30,
        5 => 31,
        6 => 30,
        7 => 31,
        8 => 31,
        9 => 30,
        10 => 31,
        11 => 30,
        12 => 31,
        _ => 30, // shouldn't happen after normalization
    }
}

/// Normalizes month and day, handling overflow/underflow.
///
/// For example, month 13 becomes month 1 of the next year.
/// Day 0 becomes the last day of the previous month.
fn normalize_date(year: i64, month: i64, day: i64) -> (i64, i64, i64) {
    // Normalize month
    let mut y = year;
    let mut m = month;

    // Handle month overflow/underflow
    if m < 1 {
        let months_back = 1 - m;
        let years_back = (months_back + 11) / 12;
        y -= years_back;
        m += years_back * 12;
    } else if m > 12 {
        let extra_months = m - 1;
        y += extra_months / 12;
        m = (extra_months % 12) + 1;
    }

    // Normalize day
    let mut d = day;
    // Handle day underflow (negative or zero days)
    while d < 1 {
        m -= 1;
        if m < 1 {
            m = 12;
            y -= 1;
        }
        d += days_in_month(y, m);
    }
    // Handle day overflow
    loop {
        let dim = days_in_month(y, m);
        if d <= dim {
            break;
        }
        d -= dim;
        m += 1;
        if m > 12 {
            m = 1;
            y += 1;
        }
    }

    (y, m, d)
}

/// Converts year/month/day to an Excel serial date number.
///
/// Returns `None` if the resulting serial would be <= 0.
fn date_to_serial(year: i64, month: i64, day: i64) -> Option<i64> {
    let (y, m, d) = normalize_date(year, month, day);

    // Count days from January 1, 1900 (serial 1)
    // We compute the number of days from a reference point.
    let mut serial: i64 = 0;

    // Add days for years 1900..(y-1)
    for yr in 1900..y {
        serial += if is_real_leap_year(yr) { 366 } else { 365 };
    }
    // If y < 1900, we need to subtract
    if y < 1900 {
        for yr in y..1900 {
            serial -= if is_real_leap_year(yr) { 366 } else { 365 };
        }
    }

    // Add days for months 1..(m-1) in year y
    for mo in 1..m {
        serial += days_in_month(y, mo);
    }

    // Add the day
    serial += d;

    // serial is now the number of days from Jan 1, 1900 (where Jan 1 = day 1)
    // But we haven't accounted for the phantom Feb 29, 1900 yet.
    // Excel serial 1 = Jan 1, 1900.
    // Our calculation gives Jan 1, 1900 = 1 day (from the d += d line when
    // year=1900, month=1, day=1: serial = 0 + 0 + 1 = 1). Good.

    // Now account for the phantom leap day:
    // Dates on or after Mar 1, 1900 (real serial >= 60 without the phantom day)
    // need to be shifted by +1 to accommodate Excel's fake Feb 29, 1900.
    // The threshold: Mar 1, 1900 in our calculation = 31 (Jan) + 28 (Feb) + 1 = 60.
    // Excel says Mar 1, 1900 = serial 61. So if our computed serial >= 60, add 1.
    if serial >= 60 {
        serial += 1;
    }

    if serial < 1 {
        None
    } else {
        Some(serial)
    }
}

/// Converts an Excel serial date number to (year, month, day).
fn serial_to_date(serial: i64) -> (i64, i64, i64) {
    // Special case: the phantom leap day
    if serial == 60 {
        return (1900, 2, 29);
    }

    // For serials 1-59 (Jan 1 - Feb 28, 1900), no adjustment needed.
    // For serials > 60, subtract 1 to remove the phantom leap day.
    let adjusted = if serial > 60 { serial - 1 } else { serial };

    // adjusted serial: 1 = Jan 1, 1900 (without phantom day)
    let mut remaining = adjusted - 1; // 0-based day count from Jan 1, 1900
    let mut year: i64 = 1900;

    // Find the year
    loop {
        let days_in_year = if is_real_leap_year(year) { 366 } else { 365 };
        if remaining < days_in_year {
            break;
        }
        remaining -= days_in_year;
        year += 1;
    }

    // Find the month
    let mut month: i64 = 1;
    loop {
        let dim = days_in_month(year, month);
        if remaining < dim {
            break;
        }
        remaining -= dim;
        month += 1;
    }

    let day = remaining + 1;
    (year, month, day)
}

/// Registers all date/time functions into the given registry.
pub fn register(registry: &mut FunctionRegistry) {
    registry.register(FunctionDef {
        name: "TODAY",
        min_args: 0,
        max_args: 0,
        param_kinds: &[],
        func: fn_today,
    });
    registry.register(FunctionDef {
        name: "NOW",
        min_args: 0,
        max_args: 0,
        param_kinds: &[],
        func: fn_now,
    });
    registry.register(FunctionDef {
        name: "DATE",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_date,
    });
    registry.register(FunctionDef {
        name: "YEAR",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_year,
    });
    registry.register(FunctionDef {
        name: "MONTH",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_month,
    });
    registry.register(FunctionDef {
        name: "DAY",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_day,
    });
    registry.register(FunctionDef {
        name: "HOUR",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_hour,
    });
    registry.register(FunctionDef {
        name: "MINUTE",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_minute,
    });
    registry.register(FunctionDef {
        name: "SECOND",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_second,
    });
    registry.register(FunctionDef {
        name: "TIME",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_time,
    });
    registry.register(FunctionDef {
        name: "DAYS",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_days,
    });
    registry.register(FunctionDef {
        name: "EDATE",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_edate,
    });
    registry.register(FunctionDef {
        name: "EOMONTH",
        min_args: 2,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_eomonth,
    });
    registry.register(FunctionDef {
        name: "WEEKDAY",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_weekday,
    });
    registry.register(FunctionDef {
        name: "WEEKNUM",
        min_args: 1,
        max_args: 2,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar],
        func: fn_weeknum,
    });
    registry.register(FunctionDef {
        name: "ISOWEEKNUM",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_isoweeknum,
    });
    registry.register(FunctionDef {
        name: "DATEVALUE",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_datevalue,
    });
    registry.register(FunctionDef {
        name: "TIMEVALUE",
        min_args: 1,
        max_args: 1,
        param_kinds: &[ParamKind::Scalar],
        func: fn_timevalue,
    });
    registry.register(FunctionDef {
        name: "DAYS360",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_days360,
    });
    registry.register(FunctionDef {
        name: "DATEDIF",
        min_args: 3,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_datedif,
    });
    registry.register(FunctionDef {
        name: "NETWORKDAYS",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Range],
        func: fn_networkdays,
    });
    registry.register(FunctionDef {
        name: "WORKDAY",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Range],
        func: fn_workday,
    });
    registry.register(FunctionDef {
        name: "YEARFRAC",
        min_args: 2,
        max_args: 3,
        param_kinds: &[ParamKind::Scalar, ParamKind::Scalar, ParamKind::Scalar],
        func: fn_yearfrac,
    });
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hour_minute_second() {
        // 0.5 = 12:00:00 PM
        let args = vec![FunctionArg::Value(CalcValue::number(0.5))];
        let ctx = EvalContext::stub();
        assert_eq!(fn_hour(&args, &ctx).as_scalar(), &ScalarValue::Number(12.0));
        assert_eq!(
            fn_minute(&args, &ctx).as_scalar(),
            &ScalarValue::Number(0.0)
        );
        assert_eq!(
            fn_second(&args, &ctx).as_scalar(),
            &ScalarValue::Number(0.0)
        );
    }

    #[test]
    fn test_time() {
        let args = vec![
            FunctionArg::Value(CalcValue::number(12.0)),
            FunctionArg::Value(CalcValue::number(30.0)),
            FunctionArg::Value(CalcValue::number(45.0)),
        ];
        let ctx = EvalContext::stub();
        let result = fn_time(&args, &ctx);
        // 12:30:45 = (12*3600 + 30*60 + 45) / 86400 = 45045 / 86400 ≈ 0.521354
        if let ScalarValue::Number(n) = result.as_scalar() {
            assert!((n - 0.521354).abs() < 0.001);
        } else {
            panic!("Expected number");
        }
    }

    #[test]
    fn test_days() {
        let args = vec![
            FunctionArg::Value(CalcValue::number(100.0)),
            FunctionArg::Value(CalcValue::number(50.0)),
        ];
        let ctx = EvalContext::stub();
        assert_eq!(fn_days(&args, &ctx).as_scalar(), &ScalarValue::Number(50.0));
    }

    #[test]
    fn test_weekday() {
        // Jan 1, 1900 = serial 1
        // Excel treats Jan 1, 1900 as Sunday (even though historically it was Monday)
        // With return_type=1 (Sun=1...Sat=7), Sunday = 1
        let args = vec![FunctionArg::Value(CalcValue::number(1.0))];
        let ctx = EvalContext::stub();
        assert_eq!(
            fn_weekday(&args, &ctx).as_scalar(),
            &ScalarValue::Number(1.0)
        ); // Sunday with type 1
    }

    #[test]
    fn test_days360() {
        let args = vec![
            FunctionArg::Value(CalcValue::number(date_to_serial(2000, 1, 1).unwrap() as f64)),
            FunctionArg::Value(CalcValue::number(
                date_to_serial(2000, 12, 31).unwrap() as f64
            )),
        ];
        let ctx = EvalContext::stub();
        assert_eq!(
            fn_days360(&args, &ctx).as_scalar(),
            &ScalarValue::Number(360.0)
        );
    }

    #[test]
    fn test_datedif_years() {
        let args = vec![
            FunctionArg::Value(CalcValue::number(date_to_serial(2000, 1, 1).unwrap() as f64)),
            FunctionArg::Value(CalcValue::number(date_to_serial(2005, 1, 1).unwrap() as f64)),
            FunctionArg::Value(CalcValue::text("Y")),
        ];
        let ctx = EvalContext::stub();
        assert_eq!(
            fn_datedif(&args, &ctx).as_scalar(),
            &ScalarValue::Number(5.0)
        );
    }

    #[test]
    fn serial_jan_1_1900() {
        assert_eq!(date_to_serial(1900, 1, 1), Some(1));
    }

    #[test]
    fn serial_feb_28_1900() {
        assert_eq!(date_to_serial(1900, 2, 28), Some(59));
    }

    #[test]
    fn serial_mar_1_1900() {
        // Mar 1, 1900 = serial 61 (serial 60 is the phantom Feb 29)
        assert_eq!(date_to_serial(1900, 3, 1), Some(61));
    }

    #[test]
    fn serial_jan_1_2000() {
        // Well-known: Jan 1, 2000 = serial 36526
        assert_eq!(date_to_serial(2000, 1, 1), Some(36526));
    }

    #[test]
    fn roundtrip_serial_to_date() {
        // Test a range of serials
        for serial in [1, 2, 31, 59, 61, 62, 100, 365, 366, 700, 36526, 44927] {
            let (y, m, d) = serial_to_date(serial);
            let back = date_to_serial(y, m, d);
            assert_eq!(
                back,
                Some(serial),
                "roundtrip failed for serial {serial}: date=({y},{m},{d})"
            );
        }
    }

    #[test]
    fn phantom_leap_day() {
        let (y, m, d) = serial_to_date(60);
        assert_eq!((y, m, d), (1900, 2, 29));
    }

    #[test]
    fn month_overflow() {
        // DATE(2000, 13, 1) should be Jan 1, 2001
        let serial_jan_2001 = date_to_serial(2001, 1, 1);
        let serial_month13 = date_to_serial(2000, 13, 1);
        assert_eq!(serial_jan_2001, serial_month13);
    }

    #[test]
    fn day_overflow() {
        // DATE(2000, 1, 32) should be Feb 1, 2000
        let serial_feb_1 = date_to_serial(2000, 2, 1);
        let serial_day32 = date_to_serial(2000, 1, 32);
        assert_eq!(serial_feb_1, serial_day32);
    }

    #[test]
    fn negative_day() {
        // DATE(2000, 2, -1) should be Jan 30, 2000
        let serial_jan_30 = date_to_serial(2000, 1, 30);
        let serial_neg = date_to_serial(2000, 2, -1);
        assert_eq!(serial_jan_30, serial_neg);
    }
}
