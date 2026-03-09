//! OOXML unit conversions for PresentationML (EMU, hundredths-of-pt, angle → CSS points/degrees).

/// Convert EMU (English Metric Units) to CSS points (1 pt = 12700 EMU).
pub fn emu_to_pt(emu: i64) -> f64 {
    emu as f64 / 12_700.0
}

/// Convert hundredths of a point to CSS points.
/// PresentationML font sizes are stored in hundredths of a point (e.g. 2400 = 24pt).
pub fn hundredths_to_pt(h: u32) -> f64 {
    h as f64 / 100.0
}

/// Convert 60000ths of a degree to CSS degrees.
/// PresentationML stores rotation as 60000ths of a degree (e.g. 5400000 = 90deg).
pub fn angle_to_degrees(a: i32) -> f64 {
    a as f64 / 60_000.0
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn emu_conversion() {
        // 1 inch = 914400 EMU = 72 pt
        assert!((emu_to_pt(914_400) - 72.0).abs() < 0.01);
        // 1 pt = 12700 EMU
        assert!((emu_to_pt(12_700) - 1.0).abs() < f64::EPSILON);
    }

    #[test]
    fn hundredths_conversion() {
        assert!((hundredths_to_pt(2400) - 24.0).abs() < f64::EPSILON);
        assert!((hundredths_to_pt(1800) - 18.0).abs() < f64::EPSILON);
    }

    #[test]
    fn angle_conversion() {
        assert!((angle_to_degrees(5_400_000) - 90.0).abs() < f64::EPSILON);
        assert!((angle_to_degrees(0) - 0.0).abs() < f64::EPSILON);
        assert!((angle_to_degrees(-5_400_000) - -90.0).abs() < f64::EPSILON);
    }
}
