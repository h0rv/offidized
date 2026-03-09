use std::path::Path;

use anyhow::Result;
use offidized_ir::{DeriveOptions, Mode};

/// Run the `derive` subcommand.
pub fn run(file: &Path, output: Option<&Path>, mode: &str, sheet: Option<&str>) -> Result<()> {
    let mode = Mode::parse_str(mode).map_err(|e| anyhow::anyhow!("{e}"))?;

    let options = DeriveOptions {
        mode,
        sheet: sheet.map(String::from),
        range: None,
    };

    let ir = offidized_ir::derive(file, options).map_err(|e| anyhow::anyhow!("{e}"))?;

    if let Some(out_path) = output {
        std::fs::write(out_path, &ir)?;
    } else {
        print!("{ir}");
    }

    Ok(())
}
