//! OOXML unit conversions (twips, half-points, EMU → CSS points).

/// Convert twips to CSS points (1 twip = 1/20 pt).
pub fn twips_to_pt(twips: u32) -> f64 {
    twips as f64 / 20.0
}

/// Convert signed twips to CSS points (used for signed indentation values).
pub fn signed_twips_to_pt(twips: i32) -> f64 {
    twips as f64 / 20.0
}

/// Convert half-points to CSS points (font sizes in OOXML are in half-points).
pub fn half_points_to_pt(hp: u16) -> f64 {
    hp as f64 / 2.0
}

/// Convert EMU (English Metric Units) to CSS points (1 pt = 12700 EMU).
pub fn emu_to_pt(emu: u32) -> f64 {
    emu as f64 / 12_700.0
}

/// Convert signed EMU to CSS points.
pub fn signed_emu_to_pt(emu: i32) -> f64 {
    emu as f64 / 12_700.0
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn twips_conversion() {
        assert!((twips_to_pt(240) - 12.0).abs() < f64::EPSILON);
        assert!((twips_to_pt(0) - 0.0).abs() < f64::EPSILON);
    }

    #[test]
    fn half_points_conversion() {
        assert!((half_points_to_pt(24) - 12.0).abs() < f64::EPSILON);
        assert!((half_points_to_pt(22) - 11.0).abs() < f64::EPSILON);
    }

    #[test]
    fn emu_conversion() {
        // 1 inch = 914400 EMU = 72 pt
        assert!((emu_to_pt(914_400) - 72.0).abs() < 0.01);
    }

    #[test]
    fn signed_conversions() {
        assert!((signed_twips_to_pt(-240) - -12.0).abs() < f64::EPSILON);
        assert!((signed_emu_to_pt(-12_700) - -1.0).abs() < f64::EPSILON);
    }
}
