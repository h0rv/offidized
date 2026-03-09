//! Financial functions: FV, IPMT, PMT.

use crate::context::EvalContext;
use crate::value::{CalcValue, XlError};

use super::{require_number, FunctionArg, FunctionDef, FunctionRegistry, ParamKind};

/// `FV(rate, nper, pmt, [pv], [type])`
///
/// Returns the future value of an investment based on periodic, constant
/// payments and a constant interest rate.
///
/// - `rate`: interest rate per period
/// - `nper`: total number of payment periods
/// - `pmt`: payment made each period (negative for outflows)
/// - `pv`: present value (default: 0)
/// - `type`: 0 = payment at end of period, 1 = payment at beginning (default: 0)
pub fn fn_fv(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let rate = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let nper = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let pmt = match require_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let pv = match args.get(3) {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => 0.0,
    };
    let payment_type = match args.get(4) {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => 0.0,
    };

    if nper == 0.0 {
        return CalcValue::number(-pv);
    }

    CalcValue::number(fv_internal(rate, nper, pmt, pv, payment_type))
}

/// Internal FV calculation.
fn fv_internal(rate: f64, nper: f64, pmt: f64, pv: f64, payment_type: f64) -> f64 {
    if rate == 0.0 {
        return -(pmt * nper + pv);
    }

    let mut pmt_adjusted = pmt;
    if payment_type != 0.0 {
        pmt_adjusted *= 1.0 + rate;
    }

    -(pmt_adjusted * ((1.0 + rate).powf(nper) - 1.0) / rate + pv * (1.0 + rate).powf(nper))
}

/// `IPMT(rate, per, nper, pv, [fv], [type])`
///
/// Returns the interest payment for a given period of an investment based on
/// periodic, constant payments and a constant interest rate.
///
/// - `rate`: interest rate per period
/// - `per`: the period for which to calculate interest (1 to nper)
/// - `nper`: total number of payment periods
/// - `pv`: present value
/// - `fv`: future value (default: 0)
/// - `type`: 0 = payment at end of period, 1 = payment at beginning (default: 0)
///
/// Returns `#NUM!` if:
/// - `nper <= 0`
/// - `rate <= -1`
/// - `per < 1` or `per > nper`
pub fn fn_ipmt(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let rate = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let per = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let nper = match require_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let pv = match require_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let fv = match args.get(4) {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => 0.0,
    };
    let payment_type = match args.get(5) {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => 0.0,
    };

    if nper <= 0.0 || rate <= -1.0 {
        return CalcValue::error(XlError::Num);
    }

    let nper_ceiling = nper.ceil();

    if per < 1.0 || per > nper_ceiling {
        return CalcValue::error(XlError::Num);
    }

    let mut ipmt = fv_internal(
        rate,
        per - 1.0,
        pmt_internal(rate, nper, pv, fv, payment_type),
        pv,
        payment_type,
    ) * rate;

    if payment_type != 0.0 {
        ipmt /= 1.0 + rate;
    }

    CalcValue::number(ipmt)
}

/// `PMT(rate, nper, pv, [fv], [type])`
///
/// Calculates the payment for a loan based on constant payments and a constant
/// interest rate.
///
/// - `rate`: interest rate per period
/// - `nper`: total number of payment periods
/// - `pv`: present value (loan amount)
/// - `fv`: future value (default: 0)
/// - `type`: 0 = payment at end of period, 1 = payment at beginning (default: 0)
///
/// Returns `#NUM!` if:
/// - `nper == 0`
/// - `rate <= -1`
pub fn fn_pmt(args: &[FunctionArg<'_>], _ctx: &EvalContext<'_>) -> CalcValue {
    let rate = match require_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let nper = match require_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let pv = match require_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let fv = match args.get(3) {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => 0.0,
    };
    let payment_type = match args.get(4) {
        Some(a) => match require_number(a) {
            Ok(v) => v,
            Err(e) => return e,
        },
        None => 0.0,
    };

    if nper == 0.0 || rate <= -1.0 {
        return CalcValue::error(XlError::Num);
    }

    CalcValue::number(pmt_internal(rate, nper, pv, fv, payment_type))
}

/// Internal PMT calculation.
fn pmt_internal(rate: f64, nper: f64, pv: f64, fv: f64, payment_type: f64) -> f64 {
    if rate == 0.0 {
        return -(pv + fv) / nper;
    }

    let timing_offset = if payment_type != 0.0 { 1.0 } else { 0.0 };

    (-fv - pv * (1.0 + rate).powf(nper))
        / (1.0 + rate * timing_offset)
        / (((1.0 + rate).powf(nper) - 1.0) / rate)
}

/// Registers all financial functions into the given registry.
pub fn register(registry: &mut FunctionRegistry) {
    registry.register(FunctionDef {
        name: "FV",
        min_args: 3,
        max_args: 5,
        param_kinds: &[
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
        ],
        func: fn_fv,
    });
    registry.register(FunctionDef {
        name: "IPMT",
        min_args: 4,
        max_args: 6,
        param_kinds: &[
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
        ],
        func: fn_ipmt,
    });
    registry.register(FunctionDef {
        name: "PMT",
        min_args: 3,
        max_args: 5,
        param_kinds: &[
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
            ParamKind::Scalar,
        ],
        func: fn_pmt,
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::value::ScalarValue;

    /// Helper to create a scalar argument.
    fn scalar_arg(v: f64) -> FunctionArg<'static> {
        FunctionArg::Value(CalcValue::number(v))
    }

    #[test]
    fn test_fv_basic() {
        let ctx = EvalContext::stub();
        // FV(5%, 10, -100, 0, 0) = $1,257.79
        let args = vec![
            scalar_arg(0.05),
            scalar_arg(10.0),
            scalar_arg(-100.0),
            scalar_arg(0.0),
            scalar_arg(0.0),
        ];
        let result = fn_fv(&args, &ctx);
        match result.as_scalar() {
            ScalarValue::Number(n) => {
                assert!((n - 1257.789254).abs() < 0.001);
            }
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_fv_zero_rate() {
        let ctx = EvalContext::stub();
        // FV(0%, 12, -100, 0, 0) = 1200
        let args = vec![
            scalar_arg(0.0),
            scalar_arg(12.0),
            scalar_arg(-100.0),
            scalar_arg(0.0),
            scalar_arg(0.0),
        ];
        let result = fn_fv(&args, &ctx);
        assert_eq!(result.as_scalar(), &ScalarValue::Number(1200.0));
    }

    #[test]
    fn test_fv_zero_nper() {
        let ctx = EvalContext::stub();
        // FV(rate, 0, pmt, -1000, 0) = 1000
        let args = vec![
            scalar_arg(0.05),
            scalar_arg(0.0),
            scalar_arg(-100.0),
            scalar_arg(-1000.0),
            scalar_arg(0.0),
        ];
        let result = fn_fv(&args, &ctx);
        assert_eq!(result.as_scalar(), &ScalarValue::Number(1000.0));
    }

    #[test]
    fn test_fv_payment_at_beginning() {
        let ctx = EvalContext::stub();
        // FV(5%, 10, -100, 0, 1) with type=1
        let args = vec![
            scalar_arg(0.05),
            scalar_arg(10.0),
            scalar_arg(-100.0),
            scalar_arg(0.0),
            scalar_arg(1.0),
        ];
        let result = fn_fv(&args, &ctx);
        match result.as_scalar() {
            ScalarValue::Number(n) => {
                // Should be approximately 1320.68 (FV with beginning payments)
                assert!((n - 1320.678717).abs() < 0.001);
            }
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_pmt_basic() {
        let ctx = EvalContext::stub();
        // PMT(1%, 12, 1000) = monthly payment for $1000 loan at 1% per month over 12 months
        let args = vec![scalar_arg(0.01), scalar_arg(12.0), scalar_arg(1000.0)];
        let result = fn_pmt(&args, &ctx);
        match result.as_scalar() {
            ScalarValue::Number(n) => {
                // Expected: -88.85 (approximately)
                assert!((n + 88.8488).abs() < 0.001);
            }
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_pmt_zero_rate() {
        let ctx = EvalContext::stub();
        // PMT(0%, 12, 1200) = -100
        let args = vec![scalar_arg(0.0), scalar_arg(12.0), scalar_arg(1200.0)];
        let result = fn_pmt(&args, &ctx);
        assert_eq!(result.as_scalar(), &ScalarValue::Number(-100.0));
    }

    #[test]
    fn test_pmt_invalid_nper_zero() {
        let ctx = EvalContext::stub();
        // PMT(0.01, 0, 1000) = #NUM!
        let args = vec![scalar_arg(0.01), scalar_arg(0.0), scalar_arg(1000.0)];
        let result = fn_pmt(&args, &ctx);
        assert_eq!(result.as_scalar(), &ScalarValue::Error(XlError::Num));
    }

    #[test]
    fn test_pmt_invalid_rate() {
        let ctx = EvalContext::stub();
        // PMT(-1.1, 12, 1000) = #NUM! (rate <= -1)
        let args = vec![scalar_arg(-1.1), scalar_arg(12.0), scalar_arg(1000.0)];
        let result = fn_pmt(&args, &ctx);
        assert_eq!(result.as_scalar(), &ScalarValue::Error(XlError::Num));
    }

    #[test]
    fn test_ipmt_basic() {
        let ctx = EvalContext::stub();
        // IPMT(1%, 1, 12, 1000) = interest payment for first period
        let args = vec![
            scalar_arg(0.01),
            scalar_arg(1.0),
            scalar_arg(12.0),
            scalar_arg(1000.0),
        ];
        let result = fn_ipmt(&args, &ctx);
        match result.as_scalar() {
            ScalarValue::Number(n) => {
                // Expected: -10.0 (1% of 1000)
                assert!((n + 10.0).abs() < 0.001);
            }
            _ => panic!("Expected number, got {:?}", result),
        }
    }

    #[test]
    fn test_ipmt_invalid_nper() {
        let ctx = EvalContext::stub();
        // IPMT(1%, 1, 0, 1000) = #NUM!
        let args = vec![
            scalar_arg(0.01),
            scalar_arg(1.0),
            scalar_arg(0.0),
            scalar_arg(1000.0),
        ];
        let result = fn_ipmt(&args, &ctx);
        assert_eq!(result.as_scalar(), &ScalarValue::Error(XlError::Num));
    }

    #[test]
    fn test_ipmt_invalid_rate() {
        let ctx = EvalContext::stub();
        // IPMT(-1.1, 1, 12, 1000) = #NUM!
        let args = vec![
            scalar_arg(-1.1),
            scalar_arg(1.0),
            scalar_arg(12.0),
            scalar_arg(1000.0),
        ];
        let result = fn_ipmt(&args, &ctx);
        assert_eq!(result.as_scalar(), &ScalarValue::Error(XlError::Num));
    }

    #[test]
    fn test_ipmt_invalid_period() {
        let ctx = EvalContext::stub();
        // IPMT(1%, 0, 12, 1000) = #NUM! (per < 1)
        let args = vec![
            scalar_arg(0.01),
            scalar_arg(0.0),
            scalar_arg(12.0),
            scalar_arg(1000.0),
        ];
        let result = fn_ipmt(&args, &ctx);
        assert_eq!(result.as_scalar(), &ScalarValue::Error(XlError::Num));

        // IPMT(1%, 13, 12, 1000) = #NUM! (per > nper)
        let args = vec![
            scalar_arg(0.01),
            scalar_arg(13.0),
            scalar_arg(12.0),
            scalar_arg(1000.0),
        ];
        let result = fn_ipmt(&args, &ctx);
        assert_eq!(result.as_scalar(), &ScalarValue::Error(XlError::Num));
    }

    #[test]
    fn test_ipmt_with_ceiling() {
        let ctx = EvalContext::stub();
        // IPMT(1%, 12, 11.5, 1000) - nper is ceiled to 12, per=12 is valid
        let args = vec![
            scalar_arg(0.01),
            scalar_arg(12.0),
            scalar_arg(11.5),
            scalar_arg(1000.0),
        ];
        let result = fn_ipmt(&args, &ctx);
        match result.as_scalar() {
            ScalarValue::Number(_) => {} // Should succeed
            _ => panic!("Expected number, got {:?}", result),
        }
    }
}
